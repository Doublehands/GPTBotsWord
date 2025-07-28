const http = require('http');
const https = require('https');
const url = require('url');
const fs = require('fs');
const path = require('path');
const net = require('net');

const PORT = 3001;
const ALLOWED_ORIGINS = [
    'http://localhost:3000',
    'https://localhost:3000',
    'http://localhost:3001',
    'https://localhost:3001',
    'null',  // 允许来自file://协议的请求
    'file://', // 允许来自file://协议的请求
    'https://localhost', // 允许来自Office加载项的请求
    'https://127.0.0.1', // 允许来自Office加载项的请求
    undefined, // 允许没有Origin的请求
    'none' // 允许特殊的Origin值
];

// 检查端口是否被占用
function isPortInUse(port) {
    return new Promise((resolve) => {
        const server = net.createServer()
            .once('error', (err) => {
                console.log(`端口 ${port} 检查错误:`, err.message);
                resolve(true);
            })
            .once('listening', () => {
                server.close();
                resolve(false);
            })
            .listen(port);
    });
}

// 创建PID文件
function createPidFile() {
    const pidFile = path.join(__dirname, 'local-server.pid');
    try {
        fs.writeFileSync(pidFile, process.pid.toString());
        console.log(`已创建PID文件: ${pidFile}`);
        
        // 进程退出时清理PID文件
        process.on('SIGINT', () => {
            try {
                fs.unlinkSync(pidFile);
                console.log('已清理PID文件');
            } catch (err) {
                console.error('清理PID文件失败:', err);
            }
            process.exit(0);
        });
        
        process.on('SIGTERM', () => {
            try {
                fs.unlinkSync(pidFile);
                console.log('已清理PID文件');
            } catch (err) {
                console.error('清理PID文件失败:', err);
            }
            process.exit(0);
        });
    } catch (err) {
        console.error('创建PID文件失败:', err);
    }
}

// 处理CORS
function handleCORS(req, res) {
    const origin = req.headers.origin;
    
    // 记录请求信息
    console.log('\n📡 收到请求:');
    console.log('- 方法:', req.method);
    console.log('- 路径:', req.url);
    console.log('- Origin:', origin);
    console.log('- Headers:', JSON.stringify(req.headers, null, 2));
    
    // 如果是预检请求，需要返回所有允许的头部
    if (req.method === 'OPTIONS') {
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, Accept, Origin, X-Requested-With');
        res.setHeader('Access-Control-Max-Age', '86400'); // 24小时
    }
    
    // 如果请求没有origin头（比如来自file://），允许所有来源
    if (!origin || origin === 'null' || origin === 'none') {
        res.setHeader('Access-Control-Allow-Origin', '*');
        console.log('✅ 允许所有来源访问');
    }
    // 如果origin在允许列表中，设置具体的origin
    else if (ALLOWED_ORIGINS.includes(origin)) {
        res.setHeader('Access-Control-Allow-Origin', origin);
        console.log(`✅ 允许来源: ${origin}`);
    }
    // 其他情况，记录日志但仍允许请求（开发环境）
    else {
        console.warn(`⚠️ 未知来源的请求: ${origin}`);
        res.setHeader('Access-Control-Allow-Origin', '*');
    }
    
    // 允许携带凭证
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    
    // 如果是OPTIONS请求，直接返回200
    if (req.method === 'OPTIONS') {
        res.writeHead(200);
        res.end();
        console.log('✅ 预检请求处理完成');
        return true;
    }
    
    return false;
}

// 代理API请求
function proxyApiRequest(req, res, apiPath) {
    console.log(`\n📡 代理请求开始: ${req.method} ${apiPath}`);
    console.log('请求头:', JSON.stringify(req.headers, null, 2));
    
    // 记录客户端IP
    const clientIP = req.headers['x-forwarded-for'] || req.connection.remoteAddress;
    console.log('客户端IP:', clientIP);
    
    const postData = [];
    
    req.on('data', chunk => {
        postData.push(chunk);
        console.log('收到数据块:', chunk.length, '字节');
    });
    
    req.on('end', () => {
        const body = Buffer.concat(postData).toString();
        console.log('\n请求体原始数据:', body);
        
        const options = {
            hostname: 'api.gptbots.ai',
            port: 443,
            path: apiPath,
            method: req.method,
            headers: {
                'Authorization': 'Bearer app-cqAuvC3vC7d7LybynuoZdK9D',
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(body),
                'User-Agent': 'GPTBots-Word-Addin/1.0',
                'Accept': 'application/json',
                'Connection': 'keep-alive'
            }
        };
        
        console.log('\n代理请求配置:', {
            url: `https://${options.hostname}${options.path}`,
            method: options.method,
            headers: options.headers
        });
        
        if (body) {
            try {
                const parsedBody = JSON.parse(body);
                console.log('\n解析后的请求体:', JSON.stringify(parsedBody, null, 2));
            } catch (e) {
                console.warn('请求体解析失败:', e.message);
                console.log('原始请求体:', body);
            }
        }
        
        const proxyReq = https.request(options, (proxyRes) => {
            console.log(`\n收到API响应: HTTP ${proxyRes.statusCode}`);
            console.log('响应头:', JSON.stringify(proxyRes.headers, null, 2));
            
            let responseData = '';
            
            proxyRes.on('data', (chunk) => {
                responseData += chunk;
                console.log('收到响应数据块:', chunk.length, '字节');
            });
            
            proxyRes.on('end', () => {
                console.log('\n完整响应数据:', responseData);
                
                try {
                    const parsedResponse = JSON.parse(responseData);
                    console.log('\n解析后的响应体:', JSON.stringify(parsedResponse, null, 2));
                } catch (e) {
                    console.warn('响应体解析失败:', e.message);
                    console.log('原始响应体:', responseData);
                }
                
                // 设置CORS头
                handleCORS(req, res);
                
                // 转发响应
                res.statusCode = proxyRes.statusCode;
                Object.entries(proxyRes.headers).forEach(([key, value]) => {
                    if (!['access-control-allow-origin'].includes(key.toLowerCase())) {
                        res.setHeader(key, value);
                        console.log(`设置响应头: ${key} = ${value}`);
                    }
                });
                
                res.end(responseData);
                console.log('\n📡 代理请求完成');
                console.log('响应状态码:', res.statusCode);
                console.log('响应大小:', responseData.length, '字节\n');
            });
        });
        
        proxyReq.on('error', (error) => {
            console.error('\n❌ 代理请求失败:', error);
            console.error('错误详情:', {
                message: error.message,
                code: error.code,
                stack: error.stack
            });
            
            handleCORS(req, res);
            res.statusCode = 500;
            const errorResponse = {
                error: '代理请求失败',
                message: error.message,
                code: error.code,
                timestamp: new Date().toISOString()
            };
            
            res.end(JSON.stringify(errorResponse));
            console.log('已发送错误响应:', errorResponse);
        });
        
        if (body) {
            proxyReq.write(body);
            console.log('已发送请求体数据');
        }
        proxyReq.end();
        console.log('代理请求已发送\n');
    });
}

async function startServer() {
    try {
        console.log('\n🚀 启动本地代理服务器...');
        
        // 检查端口是否被占用
        const portInUse = await isPortInUse(PORT);
        if (portInUse) {
            console.error(`\n❌ 错误: 端口 ${PORT} 已被占用`);
            console.log('尝试终止占用端口的进程...');
            
            // 尝试清理之前的进程
            try {
                const pidFile = path.join(__dirname, 'local-server.pid');
                if (fs.existsSync(pidFile)) {
                    const oldPid = parseInt(fs.readFileSync(pidFile, 'utf8'));
                    process.kill(oldPid, 'SIGTERM');
                    console.log(`已终止旧进程 (PID: ${oldPid})`);
                    // 等待端口释放
                    await new Promise(resolve => setTimeout(resolve, 1000));
                }
            } catch (err) {
                console.error('清理旧进程失败:', err.message);
            }
            
            // 再次检查端口
            const stillInUse = await isPortInUse(PORT);
            if (stillInUse) {
                console.error('\n❌ 无法释放端口，请手动终止占用端口的进程');
                console.log('可以使用以下命令查看占用端口的进程:');
                console.log(`  lsof -i :${PORT}`);
                console.log('然后使用以下命令终止进程:');
                console.log('  kill <PID>');
                process.exit(1);
            }
        }
        
        // 创建HTTP服务器
        const server = http.createServer((req, res) => {
            const parsedUrl = url.parse(req.url, true);
            const pathname = parsedUrl.pathname;
            
            console.log(`\n📥 收到请求: ${req.method} ${pathname}`);
            console.log('请求头:', req.headers);
            
            // 处理OPTIONS预检请求
            if (req.method === 'OPTIONS') {
                handleCORS(req, res);
                res.statusCode = 200;
                res.end();
                console.log('✅ 已处理OPTIONS预检请求\n');
                return;
            }
            
            // 代理API请求
            if (pathname.startsWith('/api/')) {
                const apiPath = pathname.replace('/api', '');
                proxyApiRequest(req, res, apiPath);
                return;
            }
            
            // 处理根路径请求
            if (pathname === '/') {
                handleCORS(req, res);
                res.setHeader('Content-Type', 'text/plain');
                res.statusCode = 200;
                res.end('本地代理服务器正在运行');
                console.log('✅ 已处理根路径请求\n');
                return;
            }
            
            // 提供静态文件服务
            let filePath = '.' + pathname;
            if (filePath === './') {
                filePath = './debug-api.html';
            }
            
            fs.readFile(filePath, (error, content) => {
                handleCORS(req, res);
                
                if (error) {
                    if (error.code === 'ENOENT') {
                        console.log(`❌ 文件未找到: ${filePath}`);
                        res.statusCode = 404;
                        res.end('文件未找到');
                    } else {
                        console.error(`❌ 服务器错误:`, error);
                        res.statusCode = 500;
                        res.end(`服务器错误: ${error.code}`);
                    }
                } else {
                    res.setHeader('Content-Type', 'text/html');
                    res.statusCode = 200;
                    res.end(content, 'utf-8');
                    console.log(`✅ 已发送文件: ${filePath}\n`);
                }
            });
        });
        
        // 创建PID文件
        createPidFile();
        
        // 启动服务器
        server.listen(PORT, () => {
            console.log('\n✅ 本地代理服务器已启动');
            console.log(`🌐 服务器地址: http://localhost:${PORT}`);
            console.log('\n可用端点:');
            console.log(`📝 创建对话: http://localhost:${PORT}/api/v1/conversation`);
            console.log(`💬 发送消息: http://localhost:${PORT}/api/v2/conversation/message`);
            console.log('\n按 Ctrl+C 停止服务器');
        });
        
        server.on('error', (error) => {
            console.error('\n❌ 服务器错误:', error.message);
            if (error.code === 'EADDRINUSE') {
                console.log(`端口 ${PORT} 已被占用，请关闭其他应用或换一个端口`);
            }
            process.exit(1);
        });
        
    } catch (error) {
        console.error('\n❌ 启动服务器失败:', error.message);
        process.exit(1);
    }
}

// 启动服务器
startServer(); 