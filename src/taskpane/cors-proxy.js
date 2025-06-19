// CORS代理解决方案

// 方案1：使用公共CORS代理
const CORS_PROXIES = [
    'https://cors-anywhere.herokuapp.com/',
    'https://api.allorigins.win/raw?url=',
    'https://corsproxy.io/?'
];

// 方案2：修改API调用以绕过CORS
window.apiWithCorsProxy = {
    async createConversation() {
        const originalUrl = 'https://api.gptbots.ai/v1/conversation';
        const proxyUrl = CORS_PROXIES[0] + originalUrl;
        
        try {
            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer app-nHIn7Ghs7maO6D3vVpnLm489',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: JSON.stringify({
                    user_id: 'word-gpt-user'
                })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            console.log('✅ 使用代理创建对话成功:', data);
            return data;
            
        } catch (error) {
            console.error('❌ 代理请求失败:', error);
            throw error;
        }
    },
    
    async sendMessage(conversationId, message) {
        const originalUrl = 'https://api.gptbots.ai/v2/conversation/message';
        const proxyUrl = CORS_PROXIES[0] + originalUrl;
        
        try {
            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer app-nHIn7Ghs7maO6D3vVpnLm489',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: JSON.stringify({
                    user_id: 'word-gpt-user',
                    conversation_id: conversationId,
                    message: message
                })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            console.log('✅ 使用代理发送消息成功:', data);
            return data;
            
        } catch (error) {
            console.error('❌ 代理消息发送失败:', error);
            throw error;
        }
    }
};

// 方案3：测试不同的代理
window.testCorsProxies = async function() {
    console.log('🔍 测试CORS代理...');
    
    for (let i = 0; i < CORS_PROXIES.length; i++) {
        const proxy = CORS_PROXIES[i];
        console.log(`测试代理 ${i + 1}: ${proxy}`);
        
        try {
            const testUrl = proxy + 'https://api.gptbots.ai/v1/conversation';
            const response = await fetch(testUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer app-nHIn7Ghs7maO6D3vVpnLm489'
                },
                body: JSON.stringify({
                    user_id: 'test-user'
                })
            });
            
            console.log(`✅ 代理 ${i + 1} 响应状态:`, response.status);
            
            if (response.ok) {
                const data = await response.json();
                console.log(`✅ 代理 ${i + 1} 可用，响应:`, data);
                return proxy; // 返回可用的代理
            }
            
        } catch (error) {
            console.log(`❌ 代理 ${i + 1} 失败:`, error.message);
        }
    }
    
    console.log('❌ 所有代理都不可用');
    return null;
};

console.log('🔧 CORS代理工具已加载');
console.log('可用命令:');
console.log('- testCorsProxies(): 测试所有CORS代理');
console.log('- apiWithCorsProxy.createConversation(): 使用代理创建对话'); 