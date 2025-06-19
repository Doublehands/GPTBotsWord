const https = require('https');

// API配置
const API_KEY = 'app-nHIn7Ghs7maO6D3vVpnLm489';
const USER_ID = 'test-user-001';

console.log('🤖 开始测试 GPTBots API...\n');

// 测试创建对话
function testCreateConversation() {
    return new Promise((resolve, reject) => {
        const data = JSON.stringify({
            user_id: USER_ID
        });

        const options = {
            hostname: 'api.gptbots.ai',
            port: 443,
            path: '/v1/conversation',
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${API_KEY}`,
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(data)
            }
        };

        console.log('📡 正在测试创建对话API...');
        console.log('请求地址:', `https://${options.hostname}${options.path}`);
        console.log('请求头:', options.headers);
        console.log('请求体:', data);
        console.log('');

        const req = https.request(options, (res) => {
            let responseData = '';

            res.on('data', (chunk) => {
                responseData += chunk;
            });

            res.on('end', () => {
                console.log('📥 创建对话响应:');
                console.log('状态码:', res.statusCode);
                console.log('响应头:', res.headers);
                console.log('响应体:', responseData);
                console.log('');

                if (res.statusCode === 200 || res.statusCode === 201) {
                    try {
                        const result = JSON.parse(responseData);
                        resolve(result);
                    } catch (e) {
                        resolve(responseData);
                    }
                } else {
                    reject(new Error(`HTTP ${res.statusCode}: ${responseData}`));
                }
            });
        });

        req.on('error', (error) => {
            console.error('❌ 创建对话请求失败:', error.message);
            reject(error);
        });

        req.write(data);
        req.end();
    });
}

// 测试发送消息
function testSendMessage(conversationId) {
    return new Promise((resolve, reject) => {
        const data = JSON.stringify({
            conversation_id: conversationId,
            response_mode: "blocking",
            messages: [
                {
                    role: "user",
                    content: "你好，请介绍一下你自己。"
                }
            ],
            conversation_config: {
                long_term_memory: false,
                short_term_memory: false
            }
        });

        const options = {
            hostname: 'api.gptbots.ai',
            port: 443,
            path: '/v2/conversation/message',
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${API_KEY}`,
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(data)
            }
        };

        console.log('📤 正在测试发送消息API...');
        console.log('请求地址:', `https://${options.hostname}${options.path}`);
        console.log('请求体:', data);
        console.log('');

        const req = https.request(options, (res) => {
            let responseData = '';

            res.on('data', (chunk) => {
                responseData += chunk;
            });

            res.on('end', () => {
                console.log('📥 发送消息响应:');
                console.log('状态码:', res.statusCode);
                console.log('响应头:', res.headers);
                console.log('响应体:', responseData);
                console.log('');

                if (res.statusCode === 200 || res.statusCode === 201) {
                    try {
                        const result = JSON.parse(responseData);
                        resolve(result);
                    } catch (e) {
                        resolve(responseData);
                    }
                } else {
                    reject(new Error(`HTTP ${res.statusCode}: ${responseData}`));
                }
            });
        });

        req.on('error', (error) => {
            console.error('❌ 发送消息请求失败:', error.message);
            reject(error);
        });

        req.write(data);
        req.end();
    });
}

// 完整测试流程
async function runFullTest() {
    try {
        console.log('='.repeat(60));
        console.log('           GPTBots API 完整测试');
        console.log('='.repeat(60));
        console.log('');

        // 步骤1：创建对话
        console.log('📋 步骤1：创建对话');
        const createResult = await testCreateConversation();
        
        let conversationId;
        if (typeof createResult === 'object' && createResult.conversation_id) {
            conversationId = createResult.conversation_id;
            console.log('✅ 创建对话成功！');
            console.log('对话ID:', conversationId);
        } else {
            throw new Error('创建对话失败或未返回conversation_id');
        }

        console.log('\n' + '-'.repeat(40) + '\n');

        // 步骤2：发送消息
        console.log('📋 步骤2：发送消息');
        const messageResult = await testSendMessage(conversationId);
        
        if (typeof messageResult === 'object' && messageResult.output && messageResult.output[0] && messageResult.output[0].content && messageResult.output[0].content.text) {
            const aiReply = messageResult.output[0].content.text;
            console.log('✅ 发送消息成功！');
            console.log('AI回复:', aiReply);
        } else {
            console.log('⚠️ 发送消息可能成功，但响应格式与预期不符');
            console.log('实际响应:', JSON.stringify(messageResult, null, 2));
        }

        console.log('\n' + '='.repeat(60));
        console.log('🎉 测试完成！API配置正确。');
        console.log('='.repeat(60));

    } catch (error) {
        console.log('\n' + '='.repeat(60));
        console.log('❌ 测试失败:', error.message);
        console.log('='.repeat(60));
        
        // 提供问题排查建议
        console.log('\n🔍 问题排查建议:');
        
        if (error.message.includes('ENOTFOUND') || error.message.includes('ECONNREFUSED')) {
            console.log('- 检查网络连接');
            console.log('- 确认API地址是否正确');
            console.log('- 检查防火墙设置');
        } else if (error.message.includes('401') || error.message.includes('403')) {
            console.log('- 检查API密钥是否正确');
            console.log('- 确认API密钥是否有足够权限');
        } else if (error.message.includes('400')) {
            console.log('- 检查请求参数格式');
            console.log('- 确认user_id格式是否正确');
        } else {
            console.log('- 查看上面的详细错误信息');
            console.log('- 确认API服务是否正常运行');
        }
    }
}

// 运行测试
runFullTest(); 