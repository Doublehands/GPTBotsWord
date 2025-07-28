// 本地代理API调用 - 复制debug-api.html中成功的方法

window.localProxyAPI = {
    // 检查本地代理是否可用
    async checkProxyAvailable() {
        const maxRetries = 3;
        const retryDelay = 1000; // 1秒
        let lastError = null;

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                console.log(`🔍 检查本地代理可用性: http://localhost:3001 (尝试 ${attempt}/${maxRetries})`);
                
                // 1. 首先测试基础连接
                const baseResponse = await fetch('http://localhost:3001/', {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json',
                        'Cache-Control': 'no-cache'
                    },
                    mode: 'cors',
                    signal: AbortSignal.timeout(5000)
                });

                if (!baseResponse.ok) {
                    throw new Error('服务未启动');
                }

                // 2. 测试API端点
                const apiResponse = await fetch('http://localhost:3001/api/v1/conversation', {
                    method: 'OPTIONS',
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                    },
                    mode: 'cors',
                    signal: AbortSignal.timeout(5000)
                });

                if (apiResponse.ok) {
                    console.log('✅ 本地代理服务器正常运行');
                    return {
                        available: true,
                        status: apiResponse.status,
                        message: '服务正常'
                    };
                } else {
                    throw new Error('服务异常');
                }
            } catch (error) {
                console.error(`❌ 检查失败 (尝试 ${attempt}/${maxRetries}):`, error);
                
                let errorMessage = '服务未启动';
                if (error.name === 'AbortError') {
                    errorMessage = '连接超时';
                } else if (error.message.includes('NetworkError')) {
                    errorMessage = '网络异常';
                }
                
                lastError = {
                    error: error,
                    message: errorMessage
                };
                
                if (attempt < maxRetries) {
                    console.log(`等待 ${retryDelay}ms 后重试...`);
                    await new Promise(resolve => setTimeout(resolve, retryDelay));
                    continue;
                }
            }
        }
        
        return {
            available: false,
            error: lastError.error,
            message: lastError.message
        };
    },
    
    // 创建对话 - 使用debug-api.html中成功的方法
    async createConversation(userId = 'MacOSJiaqi') {
        console.log('🔄 使用本地代理创建对话...');
        console.log('📍 请求 URL: http://localhost:3001/api/v1/conversation');
        console.log('📍 用户ID:', userId);
        
        try {
            const requestBody = { user_id: userId };
            console.log('📍 请求体:', requestBody);
            
            const response = await fetch('http://localhost:3001/api/v1/conversation', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                    // 注意：本地代理会自动添加Authorization头
                },
                body: JSON.stringify(requestBody),
                mode: 'cors',
                signal: AbortSignal.timeout(10000)
            });
            
            console.log('📍 响应状态:', response.status, response.statusText);
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            console.log('✅ 本地代理创建对话成功:', data);
            
            // 提取对话ID（支持多种可能的字段名）
            const conversationId = data.conversation_id || data.id || 
                                   data.data?.conversation_id || data.data?.id ||
                                   data.result?.conversation_id || data.result?.id;
            
            if (!conversationId) {
                throw new Error('响应中未找到对话ID');
            }
            
            return {
                success: true,
                conversationId: conversationId,
                data: data
            };
            
        } catch (error) {
            console.error('❌ 本地代理创建对话失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    },
    
    // 发送消息 - 使用debug-api.html中成功的方法
    async sendMessage(conversationId, message) {
        console.log('🔄 使用本地代理发送消息...');
        console.log('对话ID:', conversationId);
        console.log('消息内容:', message);
        
        try {
            // 使用debug-api.html中完全相同的请求体格式
            const requestBody = {
                conversation_id: conversationId,
                response_mode: "blocking",
                messages: [
                    {
                        role: "user",
                        content: message
                    }
                ],
                conversation_config: {
                    long_term_memory: false,
                    short_term_memory: false
                }
            };
            
            console.log('请求体:', requestBody);
            
            const response = await fetch('http://localhost:3001/api/v2/conversation/message', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                    // 注意：本地代理会自动添加Authorization头
                },
                body: JSON.stringify(requestBody)
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            console.log('✅ 本地代理发送消息成功:', data);
            
            // 提取AI回复（支持多种可能的字段名）
            let aiReply = null;
            
            // 尝试各种可能的响应格式
            if (data.answer) {
                aiReply = data.answer;
            } else if (data.message) {
                aiReply = data.message;
            } else if (data.content) {
                aiReply = data.content;
            } else if (data.response) {
                aiReply = data.response;
            } else if (data.output && data.output[0]) {
                if (data.output[0].content && data.output[0].content.text) {
                    aiReply = data.output[0].content.text;
                } else if (data.output[0].content) {
                    aiReply = data.output[0].content;
                } else if (data.output[0].text) {
                    aiReply = data.output[0].text;
                }
            } else if (data.data) {
                aiReply = data.data.answer || data.data.message || data.data.content;
            } else if (data.result) {
                aiReply = data.result.answer || data.result.message || data.result.content;
            }
            
            if (!aiReply) {
                console.warn('未能从响应中提取AI回复，使用完整响应');
                aiReply = JSON.stringify(data, null, 2);
            }
            
            return {
                success: true,
                message: aiReply,
                data: data
            };
            
        } catch (error) {
            console.error('❌ 本地代理发送消息失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    },
    
    // 完整的对话流程
    async processMessage(message, userId = 'word-gpt-user') {
        console.log('开始完整的本地代理对话流程...');
        
        try {
            // 步骤1: 创建对话
            const createResult = await this.createConversation(userId);
            if (!createResult.success) {
                throw new Error(`创建对话失败: ${createResult.error}`);
            }
            
            // 步骤2: 发送消息
            const messageResult = await this.sendMessage(createResult.conversationId, message);
            if (!messageResult.success) {
                throw new Error(`发送消息失败: ${messageResult.error}`);
            }
            
            console.log('🎉 本地代理完整流程成功！');
            return {
                success: true,
                conversationId: createResult.conversationId,
                message: messageResult.message,
                data: messageResult.data
            };
            
        } catch (error) {
            console.error('❌ 本地代理完整流程失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }
};

// 快捷测试函数
window.testLocalProxy = async function() {
    console.log('🧪 测试本地代理API...');
    
    // 检查代理可用性
    const isAvailable = await window.localProxyAPI.checkProxyAvailable();
    console.log('代理可用性:', isAvailable);
    
    if (!isAvailable.available) {
        console.log('❌ 本地代理不可用，请检查错误信息:', isAvailable.message);
        console.log('💡 运行命令: node local-server.js');
        return;
    }
    
    // 测试完整流程
    const result = await window.localProxyAPI.processMessage('你好，请介绍一下你自己。');
    
    if (result.success) {
        console.log('🎉 本地代理测试成功！');
        console.log('AI回复:', result.message);
    } else {
        console.log('❌ 本地代理测试失败:', result.error);
    }
    
    return result;
};

console.log('🔧 本地代理API已加载');
console.log('可用命令:');
console.log('- testLocalProxy(): 测试本地代理完整流程');
console.log('- localProxyAPI.checkProxyAvailable(): 检查代理可用性'); 