const http = require('http');
const https = require('https');
const url = require('url');
const fs = require('fs');
const path = require('path');

const PORT = 8080;

// 代理API请求
function proxyApiRequest(req, res, apiPath) {
    const postData = [];
    
    req.on('data', chunk => {
        postData.push(chunk);
    });
    
    req.on('end', () => {
        const body = Buffer.concat(postData).toString();
        
        const options = {
            hostname: 'api.gptbots.ai',
            port: 443,
            path: apiPath,
            method: req.method,
            headers: {
                'Authorization': 'Bearer app-nHIn7Ghs7maO6D3vVpnLm489',
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(body)
            }
        };
        
        console.log(`📡 代理请求: ${req.method} ${apiPath}`);
        console.log('请求体:', body);
        
        const proxyReq = https.request(options, (proxyRes) => {
            let responseData = '';
            
            proxyRes.on('data', (chunk) => {
                responseData += chunk;
            });
            
            proxyRes.on('end', () => {
                console.log(`📥 API响应 [${proxyRes.statusCode}]:`, responseData);
                
                // 设置CORS头
                res.setHeader('Access-Control-Allow-Origin', '*');
                res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
                res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
                res.setHeader('Content-Type', 'application/json');
                
                res.statusCode = proxyRes.statusCode;
                res.end(responseData);
            });
        });
        
        proxyReq.on('error', (error) => {
            console.error('❌ 代理请求失败:', error.message);
            res.setHeader('Access-Control-Allow-Origin', '*');
            res.statusCode = 500;
            res.end(JSON.stringify({ error: error.message }));
        });
        
        proxyReq.write(body);
        proxyReq.end();
    });
}

// 创建HTTP服务器
const server = http.createServer((req, res) => {
    const parsedUrl = url.parse(req.url, true);
    const pathname = parsedUrl.pathname;
    
    console.log(`📍 请求: ${req.method} ${pathname}`);
    
    // 处理OPTIONS预检请求
    if (req.method === 'OPTIONS') {
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
        res.statusCode = 200;
        res.end();
        return;
    }
    
    // 代理API请求
    if (pathname.startsWith('/api/')) {
        const apiPath = pathname.replace('/api', '');
        proxyApiRequest(req, res, apiPath);
        return;
    }
    
    // 提供静态文件服务
    let filePath = '.' + pathname;
    if (filePath === './') {
        filePath = './debug-api.html';
    }
    
    fs.readFile(filePath, (error, content) => {
        if (error) {
            if (error.code === 'ENOENT') {
                res.statusCode = 404;
                res.end('文件未找到');
            } else {
                res.statusCode = 500;
                res.end(`服务器错误: ${error.code}`);
            }
        } else {
            res.setHeader('Content-Type', 'text/html');
            res.statusCode = 200;
            res.end(content, 'utf-8');
        }
    });
});

server.listen(PORT, () => {
    console.log('🚀 本地服务器已启动');
    console.log(`📍 服务器地址: http://localhost:${PORT}`);
    console.log(`🔧 调试页面: http://localhost:${PORT}/debug-api.html`);
    console.log('📡 API代理路径:');
    console.log('   - 创建对话: http://localhost:8080/api/v1/conversation');
    console.log('   - 发送消息: http://localhost:8080/api/v2/conversation/message');
    console.log('');
    console.log('💡 现在可以在浏览器中测试API了！');
    console.log('按 Ctrl+C 停止服务器');
});

server.on('error', (error) => {
    console.error('❌ 服务器启动失败:', error.message);
    if (error.code === 'EADDRINUSE') {
        console.log(`端口 ${PORT} 已被占用，请关闭其他应用或换一个端口`);
    }
}); 