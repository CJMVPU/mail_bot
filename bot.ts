import 'dotenv/config';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import { extractText, getDocumentProxy } from 'unpdf';
import * as xlsx from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import { pathToFileURL } from 'url'; // 👈 新增这行


// ⭐ 无论在 Windows 还是 Linux，它都会自动找到和 bot.ts 同目录下的 台账.xlsx
const EXCEL_PATH = path.resolve(__dirname, './台账.xlsx');
// 用于记录本次运行期间已经处理过的邮件 UID，防止重复分析
const MAX_UID_CACHE_SIZE = 3000;
const processedUids = new Set<number>();

let client: ImapFlow;
// 在全局定义一个缓存
let cachedActiveOrders = new Set<string>();
// 首次启动时读入一次
cachedActiveOrders = getActiveOrders(EXCEL_PATH);
// --- 优雅退出机制 ---
// ⭐ 新增全局状态锁，防止关机时触发重连风暴
let isShuttingDown = false;


// 监听 Excel 文件的修改，一旦修改，立刻在后台静默重载
fs.watch(EXCEL_PATH, (eventType) => {
    if (eventType === 'change') {
        console.log('🔄 检测到 Excel 台账更新，正在重新加载数据...');
        const newOrders = getActiveOrders(EXCEL_PATH);
        // ⭐ 防御机制：只有成功读取到数据时，才替换旧缓存。
        // 防止 Excel 保存瞬间的文件锁导致读取为空，把已有缓存洗掉
        if (newOrders.size > 0) {
            cachedActiveOrders = newOrders;
            console.log('✅ 缓存已成功热更新！');
        } else {
            console.log('⚠️ 读取为空或文件被锁，保持使用旧缓存。');
        }
    }
});

// --- 数据读取模块 ---
/**
 * 从本地 Excel 台账中读取 WorkingShips 表里活跃的 MBL、HBL 和 柜号
 * @param excelFilePath 你的台账.xlsx文件的绝对路径
 * @returns Set<string> 使用 Set 代替 Array，将后续的查找性能提升至 O(1)
 */
function getActiveOrders(excelFilePath: string = EXCEL_PATH): Set<string> {
    const activeOrders = new Set<string>(); 

    try {
        if (!fs.existsSync(excelFilePath)) {
            console.error(`❌ 找不到 Excel 文件: ${excelFilePath}`);
            return activeOrders;
        }

        const workbook = xlsx.readFile(excelFilePath);
        const targetSheetName = 'WorkingShips';
        const worksheet = workbook.Sheets[targetSheetName];

        if (!worksheet) {
            console.error(`❌ 在 Excel 中找不到名为 '${targetSheetName}' 的工作表`);
            return activeOrders;
        }
            
        const rows = xlsx.utils.sheet_to_json<Record<string, any>>(worksheet, { defval: '' });

        // ✅ 内存优化 + 脏数据多单号粉碎机
        const cleanAndAdd = (val: any) => {
            if (val && typeof val === 'string') {
                // 利用正则同时支持劈开：斜杠 /、反斜杠 \、逗号 ,、中文逗号 ，、顿号 、、分号 ; 以及 换行符
                const parts = val.split(/[/\\[\]\n,;，、]/);
                
                for (const part of parts) {
                    // 对劈开后的每一个子串进行标准的去空、去横杠、转大写操作
                    const cleanVal = part.replace(/[\s-]/g, '').toUpperCase();
                    if (cleanVal) {
                        activeOrders.add(cleanVal);
                    }
                }
            } else if (val && typeof val === 'number') {
                activeOrders.add(val.toString());
            }
        };

        for (const row of rows) {
            const mbl = row['MBL'] || row['mbl'] || '';
            const hbl = row['HBL'] || row['hbl'] || '';
            const containerNo = row['柜号'] || row['柜 号'] || '';

            cleanAndAdd(mbl);
            cleanAndAdd(hbl);
            cleanAndAdd(containerNo);
        }

        console.log(`✅ 成功从 WorkingShips 表中提取了 ${activeOrders.size} 个唯一的单号/柜号。`);
        return activeOrders;

    } catch (error) {
        console.error("读取 Excel 文件时发生错误:", error);
        return activeOrders;
    }
}

// filter: regex
function isMailRelevant(text: string, cachedActiveOrders: Set<string>): boolean {
  const containerRegex = /[A-Z]{4}\s*-?\s*\d{7}/gi;
  // 1. 放宽提单号的正则下限。从 8 降到 6，确保被截断的纯数字单号也能被抓取
  const billRegex = /[A-Z0-9]{6,20}/gi; 

  const foundContainers = text.match(containerRegex) || [];
  const foundBills = text.match(billRegex) || [];

  const extractedTokens = [...foundContainers, ...foundBills].map(
    token => token.replace(/[\s-]/g, '').toUpperCase()
  );

  for (const token of extractedTokens) {
    // 方案 A：O(1) 瞬间精确匹配 (针对标准情况)
    if (cachedActiveOrders.has(token)) {
      console.log(`命中单号: ${token}`);
      return true;
    }

    // 方案 B：船司“掐头”容错匹配 (针对变态情况)
    // 为了防止把年份（如 2026）这种短数字误认作单号，我们设定一个安全阈值，长度 >= 6 才进行模糊匹配
    if (token.length >= 6) {
      for (const order of cachedActiveOrders) {
        // 双向容错：
        // 情况 1: 台账存的是 COSU123456，邮件里写的是 123456 (order.endsWith(token))
        // 情况 2: 台账存的是 123456，邮件里写的是 COSU123456 (token.endsWith(order))
        if (order.endsWith(token) || token.endsWith(order)) {
          console.log(`[HIT] 模糊容错匹配: 邮件中的 "${token}" 对应台账单号 "${order}"`);
          return true;
        }
      }
    }
  }
  return false;
}

// --- 核心处理模块 ---
async function processUnseenMessage(): Promise<void> {
    const unseenUids = await client.search({ seen: false }) as number[];
    
    if (unseenUids.length === 0) {
        return; 
    }

    if (cachedActiveOrders.size === 0) {
        console.log("⚠️ 没有获取到活跃订单数据，跳过本次过滤。");
        return;
    }

    const uidsToMarkRead: number[] = [];

    // 1. 读取阶段
    for await (let message of client.fetch(unseenUids, { source: true, uid: true })) {
        // ⭐ 新增：如果已经拉闸，立刻停止处理剩下的邮件！
        if (isShuttingDown) {
            console.log('⚠️ 系统正在关闭，中止剩余邮件读取...');
            break;
        }

        try {
            // ⭐ 新增：如果这个 UID 之前处理过了，直接跳过！
            if (processedUids.has(message.uid)) continue;
            // ⭐ 新增：立刻把当前 UID 记入脑海
            processedUids.add(message.uid);

            // 👇 新增的滑动窗口防内存泄漏逻辑
            if (processedUids.size > MAX_UID_CACHE_SIZE) {
                // 因为 JS 的 Set 保留了插入顺序，所以 values().next() 永远拿到的是最老、最早插进来的那个 UID
                const oldestUid = processedUids.values().next().value!;
                processedUids.delete(oldestUid);
            }

            if (!message.source) continue;

            const parsed = await simpleParser(message.source);
            const subject = parsed.subject || '无主题';
            let emailContent = `${subject}\n${parsed.text || ''}`;

            console.log(`\n正在分析: "${subject}"`);
            
            // ⭐ 新增：PDF 附件“透视”逻辑 ⭐
            if (parsed.attachments && parsed.attachments.length > 0) {
                for (const attachment of parsed.attachments) {
                    // 只处理 PDF 格式的附件
                    if (attachment.contentType === 'application/pdf') {
                        console.log(`📎 发现 PDF 附件: [${attachment.filename}]，正在提取文本...`);
                        try {
                            // 1. 获取 Linux 的绝对路径
                            const cMapDir = path.join(__dirname, 'node_modules/pdfjs-dist/cmaps/');
                            const fontDir = path.join(__dirname, 'node_modules/pdfjs-dist/standard_fonts/');

                            // ⭐ 2. 终极魔法：将绝对路径转换为标准的 file:// URL 协议，并在末尾强制加上斜杠！
                            const cMapUrl = pathToFileURL(cMapDir).href + '/';
                            const standardFontDataUrl = pathToFileURL(fontDir).href + '/';

                            const pdf = await getDocumentProxy(
                                new Uint8Array(attachment.content as Buffer),
                                {
                                    // 👈 终极静音：我们不要外挂字体包了，直接让它用内置的纯文本映射引擎！
                                    // cMapUrl: cMapUrl,
                                    // cMapPacked: true,
                                    // standardFontDataUrl: standardFontDataUrl,
                                    
                                    // 强制关闭渲染过程中的所有图形/字体缺失警告
                                    verbosity: 0,
                                    useSystemFonts: true 
                                } as any
                            );
                            
                            // mergePages: true 会自动把多页 PDF 的纯文本完美拼接在一起
                            const { text } = await extractText(pdf, { mergePages: true });
                            emailContent += `\n--- PDF Content (${attachment.filename}) ---\n${text}`;
                            console.log(`   └─ 提取成功，内容已合并。`);
                        } catch (pdfErr) {
                            console.error(`   └─ ❌ 解析 PDF 附件失败:`, pdfErr);
                        }
                    }
                }
            }
            const relevant = isMailRelevant(emailContent, cachedActiveOrders);

            if (!relevant) {
                uidsToMarkRead.push(message.uid);
                console.log(`[加入过滤队列] 无关邮件 🗑️`);
            } else {
                console.log(`[保留] 订单相关，保留在收件箱中 🔔`);
            }
        } catch (err) {
            console.error(`解析邮件 UID ${message.uid} 时出错:`, err);
        }
    }

    // 2. 写入阶段
    if (uidsToMarkRead.length > 0) {
        console.log(`\n正在批量将 ${uidsToMarkRead.length} 封无关邮件标记为已读...`);
        await client.messageFlagsAdd(uidsToMarkRead, ['\\Seen'], { uid: true });
        console.log(`✅ 批量标记完成！`);
    }
}

// --- 并发控制锁 ---
let isProcessing = false;     // 当前是否正在处理邮件？
let hasPendingMail = false;   // 处理期间，是否有新邮件排队？

// 专门用于替代直接调用 processUnseenMessage 的安全触发器
async function safeTriggerProcess() {
    // 1. 如果当前已经有线程在干活了，不要并发启动！
    if (isProcessing) {
        console.log('⏳ [并发拦截] 发现新邮件到达，当前正在忙碌，已加入待处理队列...');
        hasPendingMail = true; // 打个“欠条”，让当前线程干完活之后再跑一圈
        return;
    }

    // 2. 加锁
    isProcessing = true;

    try {
        // 3. 核心循环：只要有“欠条”，就一直跑，直到把积压的新邮件全部消化完
        do {
            hasPendingMail = false; // 撕毁现有的欠条
            await processUnseenMessage();
        } while (hasPendingMail); 
    } catch (err) {
        console.error('❌ [队列异常] 处理邮件队列时发生错误:', err);
    } finally {
        // 4. 彻底干完活了，释放锁
        isProcessing = false;
    }
}

// --- 主程序生命周期 ---
async function startBot() {
    // ⚠️ 不用 await client.connect()，因为外层引擎已经连过了
    let lock;
    try {
        lock = await client.getMailboxLock('INBOX');
        console.log("INBOX locked, monitoring...");

        // 2. 挂载异步监听锁
        client.on('exists', () => {
            console.log(`\n📨 detected new email, analyzing...`);
            safeTriggerProcess();
        });

        // 1. 启动时主动查漏补缺（首次全面扫描）
        safeTriggerProcess();

        // 3. ⭐ 核心防线：用 Promise 强行阻塞住当前线程！
        // 它会一直挂起在这里听指挥。只要连接不断开，它就永远不往下走。
        // 一旦底层触发 close 或 error，立刻 reject 抛出异常，触发外层引擎的指数退避重连。
        await new Promise((_, reject) => {
            client.once('close', () => reject(new Error('IMAP connection closed unexpectedly')));
            client.once('error', (err: Error) => reject(err));
        });
    } catch (error) {
        // 如果上面 Promise 抛出了断网异常，会被这里捕获，然后再扔给外层的重连引擎
        console.error("BOT runtime error / disconnected:", (error as Error).message);
        throw error; 
    } finally {
        // ⭐ 修复：只有当客户端还活着，并且可用时，才去释放锁！
        if (lock && client && client.usable) {
            try { 
                lock.release(); 
                console.log("INBOX lock released.");
            } catch (e) {
                // 静默忽略释放锁时可能产生的残余报错
            }
        }
    }
}

// --- 核心重连引擎 (指数退避算法) ---
async function connectWithBackoff() {
    let attempt = 0;
    const baseDelay = 3000;  // 基础等待时间：3秒
    const maxDelay = 300000; // 最大封顶时间：5分钟 (300000毫秒)

    while (true) {
        try {
            console.log(`\n🔄 [系统] 尝试连接 IMAP 服务器 (第 ${attempt + 1} 次) ...`);

            // ⚠️ 核心：每次重连必须 new 一个全新实例，防止底层 Socket 状态脏乱
            client = new ImapFlow({
                host: 'c2.icoremail.net', // 记得换成你的服务器
                port: 993,
                secure: true,
                auth: {
                    user: process.env.EMAIL_USER as string,
                    pass: process.env.EMAIL_PASS as string
                },
                logger: false
            });

            // 仅仅记录日志，坚决不 process.exit()
            // 解决 TS 报错：显式声明 err 为 Error 类型
            client.on('error', (err: Error) => {
                console.error('\n💥 [底层拦截] IMAP 发生异常:', err.message);
            });
            // 监听底层意外断开。由于 ImapFlow 内部机制，断开时会触发 close，
            // 此时当前的 await client.idle() 会抛出异常，从而被外层的 catch 捕获。
            client.on('close', () => {
                console.log('🔌 [警告] IMAP 底层连接已断开！');
            });

            await client.connect();
            console.log('✅ [系统] IMAP 连接成功！');

            // 一旦连接成功且没报错，立刻将退避重试次数清零
            attempt = 0;

            // 启动你的主监听逻辑 (如果 startBot 内部出错抛出异常，会跳到 catch)
            await startBot();

        } catch (err) {
            // ⭐ 新增防御：如果已经拉下总电闸，直接跳出死循环，让程序安静死去
            if (isShuttingDown) {
                console.log('👋 系统正在安全关闭，终止退避重连机制。');
                return;
            }
            
            console.error('\n💥 [异常] 运行中断或连接失败:', (err as Error).message);
            
            // 安全清理：不管实例死没死透，强制登出并释放内存
            if (client && client.usable) {
                try { await client.logout(); } catch (e) {}
            }

            // 核心算法：基础时间 * (2 的 attempt 次方)
            // Math.min 确保重试间隔不会无限增长，最高卡在 maxDelay (5分钟)
            const delay = Math.min(baseDelay * Math.pow(2, attempt), maxDelay);
            
            console.log(`⏳ [重连] 等待 ${delay / 1000} 秒后执行下一次退避重连...`);
            
            // 阻塞当前循环，等待 delay 毫秒
            await new Promise(resolve => setTimeout(resolve, delay));
            
            // 增加失败次数
            attempt++;
        }
    }
}

// 拦截 Ctrl+C 的代码稍微调整一下，确保能退出这个死循环
process.on('SIGINT', async () => {
    if (isShuttingDown) return; // 防止狂按 Ctrl+C
    isShuttingDown = true;      // 拉下总电闸！

    console.log('\n🛑 接收到退出信号，正在安全清理...');

    // ⭐ 新增：3秒倒计时！不管底层在干什么，3秒后绝对强制杀死进程
    setTimeout(() => {
        console.log('⏳ 清理超时，强制终止进程。');
        process.exit(0);
    }, 3000);

    if (client && client.usable) {
        try { await client.logout(); } catch (e) {}
    }
    process.exit(0); 
});

// 引擎点火！
connectWithBackoff();

