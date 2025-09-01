/*
 * API配置文件
 * 已为GPTBots API进行配置
 */

// API配置对象
const API_CONFIG = {
    // GPTBots API基础URL
    baseUrl: 'https://api.gptbots.ai',
    
    // 创建对话端点
    createConversationEndpoint: '/v1/conversation',
    
    // 发送消息端点
    chatEndpoint: '/v2/conversation/message',
    
    // 请求超时时间 (毫秒)
    timeout: 30000,
    
    // 请求头配置
    headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer app-nHIn7Ghs7maO6D3vVpnLm489' // 您的API密钥
    },
    
    // 默认请求参数
    defaultParams: {
        response_mode: 'blocking',
        conversation_config: {
            long_term_memory: false,
            short_term_memory: false
        }
    },
    
    // 用户ID（GPTBots需要）
    userId: 'word-gpt-plus-user', // 您可以自定义这个ID
    
    // API响应格式映射（根据官方文档确认）
    responseMapping: {
        // 创建对话响应中的对话ID字段
        conversationId: 'conversation_id',
        // 消息响应中的内容字段（在output[0].content.text路径下）
        message: 'output[0].content.text',
        // 错误信息字段
        error: 'message',
        // 状态字段
        status: 'code'
    }
};

// GPTBots API专用预设配置
const API_PRESETS = {
    // GPTBots格式
    gptbots: {
        baseUrl: 'https://api.gptbots.ai',
        createConversationEndpoint: '/v1/conversation',
        chatEndpoint: '/v2/conversation/message',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer app-nHIn7Ghs7maO6D3vVpnLm489'
        },
        defaultParams: {
            response_mode: 'blocking',
            conversation_config: {
                long_term_memory: false,
                short_term_memory: false
            }
        },
        responseMapping: {
            conversationId: 'conversation_id',
            message: 'output[0].content.text',
            error: 'message',
            status: 'code'
        }
    },
    
    // OpenAI格式（保留作为备选）
    openai: {
        baseUrl: 'https://api.openai.com',
        chatEndpoint: '/v1/chat/completions',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer YOUR_OPENAI_API_KEY'
        },
        defaultParams: {
            model: 'gpt-3.5-turbo',
            temperature: 0.7,
            max_tokens: 2000,
        },
        responseMapping: {
            message: 'choices[0].message.content',
            error: 'error.message',
            status: 'error.code'
        }
    }
};

// 辅助函数：应用预设配置
function applyPreset(presetName) {
    if (API_PRESETS[presetName]) {
        const preset = API_PRESETS[presetName];
        Object.assign(API_CONFIG, preset);
        console.log(`已应用 ${presetName} 预设配置`);
    } else {
        console.warn(`未找到预设配置: ${presetName}`);
    }
}

// 辅助函数：获取嵌套对象属性值
function getNestedValue(obj, path) {
    return path.split('.').reduce((current, key) => {
        if (key.includes('[') && key.includes(']')) {
            const arrayKey = key.substring(0, key.indexOf('['));
            const index = parseInt(key.substring(key.indexOf('[') + 1, key.indexOf(']')));
            return current && current[arrayKey] && current[arrayKey][index];
        }
        return current && current[key];
    }, obj);
}

// 辅助函数：构建创建对话的URL
function getCreateConversationUrl() {
    return `${API_CONFIG.baseUrl}${API_CONFIG.createConversationEndpoint}`;
}

// 辅助函数：构建发送消息的URL
function getChatUrl() {
    return `${API_CONFIG.baseUrl}${API_CONFIG.chatEndpoint}`;
}

// 辅助函数：构建创建对话的请求数据
function buildCreateConversationData() {
    return {
        user_id: API_CONFIG.userId
    };
}

// 辅助函数：构建发送消息的请求数据
function buildChatRequestData(conversationId, messages, customParams = {}) {
    return {
        conversation_id: conversationId,
        messages: messages,
        ...API_CONFIG.defaultParams,
        ...customParams
    };
}

// 辅助函数：解析创建对话的响应
function parseCreateConversationResponse(response) {
    try {
        const conversationIdField = API_CONFIG.responseMapping.conversationId;
        const errorField = API_CONFIG.responseMapping.error;
        
        let conversationId = getNestedValue(response, conversationIdField);
        if (!conversationId) {
            conversationId = response.conversation_id || response.id;
        }
        
        let error = getNestedValue(response, errorField);
        if (!error) {
            error = response.error || response.message;
        }
        
        return {
            conversationId: conversationId,
            error: error,
            success: !!conversationId && !error
        };
    } catch (e) {
        console.error('解析创建对话响应失败:', e);
        return {
            conversationId: null,
            error: '响应解析失败',
            success: false
        };
    }
}

// 辅助函数：解析消息响应
function parseChatResponse(response) {
    try {
        const messageField = API_CONFIG.responseMapping.message;
        const errorField = API_CONFIG.responseMapping.error;
        
        let message = getNestedValue(response, messageField);
        
        // 如果没有找到消息，尝试GPTBots的常见路径
        if (!message && response.output && response.output.length > 0) {
            // 尝试第一个输出组件的文本内容
            message = response.output[0]?.content?.text;
        }
        
        // 继续尝试其他可能的字段
        if (!message) {
            message = response.answer || response.message || response.content || response.response || response.text;
        }
        
        let error = getNestedValue(response, errorField);
        if (!error) {
            error = response.error || response.message || response.detail;
        }
        
        return {
            message: message,
            error: error,
            success: !!message && !error
        };
    } catch (e) {
        console.error('解析消息响应失败:', e);
        return {
            message: null,
            error: '响应解析失败',
            success: false
        };
    }
}

// 导出配置和辅助函数
if (typeof module !== 'undefined' && module.exports) {
    // Node.js环境
    module.exports = {
        API_CONFIG,
        API_PRESETS,
        applyPreset,
        getCreateConversationUrl,
        getChatUrl,
        buildCreateConversationData,
        buildChatRequestData,
        parseCreateConversationResponse,
        parseChatResponse
    };
} else {
    // 浏览器环境
    window.API_CONFIG = API_CONFIG;
    window.API_PRESETS = API_PRESETS;
    window.applyPreset = applyPreset;
    window.getCreateConversationUrl = getCreateConversationUrl;
    window.getChatUrl = getChatUrl;
    window.buildCreateConversationData = buildCreateConversationData;
    window.buildChatRequestData = buildChatRequestData;
    window.parseCreateConversationResponse = parseCreateConversationResponse;
    window.parseChatResponse = parseChatResponse;
}

/*
 * GPTBots API 配置说明：
 * 
 * 1. API密钥已设置：app-nHIn7Ghs7maO6D3vVpnLm489
 * 2. 支持两步API调用：
 *    - 第一步：创建对话 (POST /v1/conversation)
 *      响应格式：{"conversation_id": "657303a8a764d47094874bbe"}
 *    - 第二步：发送消息 (POST /v2/conversation/message)
 *      响应格式：{"output": [{"content": {"text": "AI回复内容"}}]}
 * 3. 响应格式已根据官方文档配置：
 *    - 对话ID: conversation_id
 *    - AI回复: output[0].content.text
 * 4. 支持的参数：
 *    - response_mode: "blocking" (阻塞式响应)
 *    - conversation_config: 对话配置选项
 * 5. 如需修改配置，请编辑上面的API_CONFIG对象
 */

// 直接API调用函数
window.directAPI = {
    // 创建对话
    async createConversation() {
        console.log('🔄 创建新对话...');
        
        try {
            const response = await fetch(getCreateConversationUrl(), {
                method: 'POST',
                headers: API_CONFIG.headers,
                body: JSON.stringify(buildCreateConversationData())
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            console.log('✅ 对话创建成功:', data);
            
            const parsed = parseCreateConversationResponse(data);
            if (!parsed.success) {
                throw new Error(parsed.error || '创建对话失败');
            }
            
            return parsed.conversationId;
            
        } catch (error) {
            console.error('❌ 创建对话失败:', error);
            throw error;
        }
    },
    
    // 发送消息
    async sendMessage(conversationId, message) {
        console.log('🔄 发送消息...');
        
        try {
            const messages = [
                {
                    role: 'user',
                    content: message
                }
            ];
            
            const requestData = buildChatRequestData(conversationId, messages);
            
            const response = await fetch(getChatUrl(), {
                method: 'POST',
                headers: API_CONFIG.headers,
                body: JSON.stringify(requestData)
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            console.log('✅ 消息发送成功:', data);
            
            const parsed = parseChatResponse(data);
            if (!parsed.success) {
                throw new Error(parsed.error || '消息处理失败');
            }
            
            return parsed.message;
            
        } catch (error) {
            console.error('❌ 发送消息失败:', error);
            throw error;
        }
    },
    
    // 完整的对话流程
    async processMessage(message) {
        console.log('🚀 开始处理消息:', message);
        
        try {
            // 创建对话
            const conversationId = await this.createConversation();
            
            // 发送消息
            const reply = await this.sendMessage(conversationId, message);
            
            console.log('🎉 消息处理完成');
            return reply;
            
        } catch (error) {
            console.error('❌ 消息处理失败:', error);
            throw error;
        }
    }
}; 