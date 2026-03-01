import 'dotenv/config';
import { GoogleGenerativeAI, SchemaType, Tool } from '@google/generative-ai';
import { exec } from 'child_process';
import { promisify } from 'util';


// 🌟 1. 智能环境嗅探：如果不是生产环境，强行启用本地代理接管
import { setGlobalDispatcher, ProxyAgent } from 'undici';
if (process.env.NODE_ENV !== 'production') {
    // 本地开发时，走你的 Windows 镜像翻墙隧道，绝不丢包！
    const dispatcher = new ProxyAgent('http://127.0.0.1:33564');
    setGlobalDispatcher(dispatcher);
    console.log(`🌐 [系统-开发模式] 已强行接管 Node 底层网络，走本地代理: 33564`);
}

// 把回调形式的 exec 包装成 Promise，方便我们用 async/await 优雅调用
const execAsync = promisify(exec);

// --- 1. 定义本地执行的真实函数 (你的“手”和“脚”) ---
const localFunctions: Record<string, Function> = {
    // 脚本 A：启动邮件过滤机器人
    startMailBot: async () => {
        console.log("🚀 [本地执行] 正在后台启动邮件过滤机器人...");
        try {
            // 这里用 PM2 启动我们刚才写的 bot.ts
            await execAsync('pm2 start bot.ts --name "mail-bot" --interpreter node --node-args="--import tsx"');
            return { status: "success", message: "邮件机器人已成功在后台启动并运行。" };
        } catch (error: any) {
            return { status: "error", message: `启动失败: ${error.message}` };
        }
    },

    // 👇 新增的脚本 B：停止并清理
    stopMailBot: async () => {
        console.log("🛑 [本地执行] 正在停止邮件过滤机器人...");
        try {
            // 先停止进程，再从 pm2 列表中彻底清理掉
            await execAsync('pm2 stop mail-bot && pm2 delete mail-bot');
            return { status: "success", message: "邮件机器人已成功停止并从后台移除。" };
        } catch (error: any) {
            // 如果进程本来就不存在，pm2 可能会报错，我们优雅地把状态告诉 AI
            return { status: "error", message: `停止失败（可能机器人并未运行）: ${error.message}` };
        }
    },

    // 脚本 B：调用外部 Python 脚本查台账 (假设你未来有一个 python 脚本)
    queryLogistics: async ({ orderNo }: { orderNo: string }) => {
        console.log(`🔍 [本地执行] 正在查询单号/柜号: ${orderNo}`);
        // 模拟执行耗时操作或者调用 python 脚本: execAsync(`python search.py ${orderNo}`)
        return { 
            status: "success", 
            orderNo: orderNo, 
            data: "已清关，预计明天送达LAX仓库。" 
        };
    }
};

// --- 2. 向 AI 注册你的工具说明书 (你的“大脑”路由表) ---
const toolsConfig:Tool[] = [{
    functionDeclarations: [
        {
            name: "startMailBot",
            description: "启动或唤醒本地的 IMAP 物流邮件过滤机器人服务。",
        },
        // 👇 新增的 AI 认知说明书
        {
            name: "stopMailBot",
            description: "停止、关闭并销毁本地正在后台运行的 IMAP 物流邮件过滤机器人服务。",
        },
        {
            name: "queryLogistics",
            description: "在本地 Excel 台账或数据库中查询特定柜号或提单号的最新物流状态。",
            parameters: {
                type: SchemaType.OBJECT,
                properties: {
                    orderNo: { type: SchemaType.STRING, description: "需要查询的提单号或柜号，例如 TCNU1234567" }
                },
                required: ["orderNo"]
            }
        }
    ]
}];

// --- 3. 核心调度引擎 ---
async function runAgent(userPrompt: string) {
    console.log(`\n👨‍💻 你指令: "${userPrompt}"`);
    console.log(`🤖 AI 正在思考如何调度...`);

    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY as string);
    // 🌟 2. 动态路由基地址
    // 本地开发不配 baseUrl（默认走代理连官方），线上生产环境才会使用你的免翻墙域名
    const baseUrl = process.env.NODE_ENV === 'production' 
        ? "https://edgerouting.uk" 
        : undefined;
        
    const model = genAI.getGenerativeModel(
        { model: "gemini-3-flash-preview", tools: toolsConfig },
        baseUrl ? { baseUrl } : undefined 
    );

    const chat = model.startChat();
    const result = await chat.sendMessage(userPrompt);
    const response = result.response;

    // 🌟 核心魔法：拦截 AI 的函数调用请求
    const functionCalls = response.functionCalls();

    if (functionCalls && functionCalls.length > 0) {
        // 遍历 AI 决定调用的所有函数
        for (const call of functionCalls) {
            console.log(`\n⚡ [拦截到 AI 指令] 准备调用本地函数: ${call.name}`);
            console.log(`📦 提取的参数:`, call.args);

            // 1. 在本地真实执行对应的 TypeScript/Python 脚本
            const targetFunction = localFunctions[call.name];
            let localResult;

            if (targetFunction) {
                localResult = await targetFunction(call.args);
                console.log(`✅ [本地执行完毕] 结果:`, localResult);
            } else {
                localResult = { error: "本地未找到该指令对应的脚本" };
            }

            // 2. 将本地脚本的执行结果“喂”回给 AI，让它生成人类看得懂的总结
            console.log(`💬 正在将脚本结果汇报给 AI...`);
            const finalResult = await chat.sendMessage([{
                functionResponse: {
                    name: call.name,
                    response: localResult
                }
            }]);

            console.log(`\n🤖 管家回复:\n${finalResult.response.text()}`);
        }
    } else {
        // 如果用户的指令不需要调用任何脚本，AI 就直接当普通聊天回复
        console.log(`\n🤖 管家回复:\n${response.text()}`);
    }
}

// --- 🎯 测试一下 ---
// 你可以随时修改这句话，看看 AI 会怎么动态路由！
runAgent("帮我启动一下邮件过滤机器人");
