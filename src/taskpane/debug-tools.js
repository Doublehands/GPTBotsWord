// 调试工具集合
window.debugTools = {
    // 测试基础功能
    testBasicFlow: async function() {
        console.log('🧪 开始测试基础功能流程...');
        
        try {
            // 1. 测试Word连接
            console.log('1️⃣ 测试Word连接...');
            const wordTest = await this.testWordConnection();
            if (!wordTest.success) {
                throw new Error(`Word连接失败: ${wordTest.error}`);
            }
            console.log('✅ Word连接正常');
            
            // 2. 测试内容获取
            console.log('2️⃣ 测试内容获取...');
            const content = await this.getWordContent();
            if (!content || content.length === 0) {
                throw new Error('无法获取Word内容，请确保选中了文本');
            }
            console.log(`✅ 内容获取成功: ${content.length} 个字符`);
            console.log('内容预览:', content.substring(0, 100));
            
            // 3. 测试结果显示
            console.log('3️⃣ 测试结果显示...');
            const testResult = `测试结果 - 当前时间: ${new Date().toLocaleString()}\n\n原始内容: ${content}\n\n这是一个测试结果，用于验证结果显示功能是否正常工作。`;
            this.displayResult(testResult);
            console.log('✅ 结果显示成功');
            
            // 4. 测试按钮状态
            console.log('4️⃣ 测试按钮状态...');
            this.testButtonStates();
            
            console.log('🎉 基础功能测试完成！所有功能正常工作。');
            return { success: true, message: '基础功能测试通过' };
            
        } catch (error) {
            console.error('❌ 基础功能测试失败:', error);
            return { success: false, error: error.message };
        }
    },
    
    // 测试Word连接
    testWordConnection: function() {
        return new Promise((resolve) => {
            Word.run(async (context) => {
                try {
                    // 测试获取文档信息
                    const doc = context.document;
                    doc.load('title');
                    
                    // 测试获取选中内容
                    const selection = context.document.getSelection();
                    selection.load('text');
                    
                    // 测试获取文档内容
                    const body = context.document.body;
                    body.load('text');
                    
                    await context.sync();
                    
                    console.log('Word连接测试结果:');
                    console.log('- 文档标题:', doc.title || '无标题');
                    console.log('- 选中文本:', selection.text || '无选中');
                    console.log('- 文档总长度:', body.text.length);
                    
                    resolve({ 
                        success: true, 
                        selection: selection.text,
                        documentLength: body.text.length 
                    });
                    
                } catch (error) {
                    console.error('Word连接失败:', error);
                    resolve({ success: false, error: error.message });
                }
            });
        });
    },
    
    // 获取Word内容
    getWordContent: function(source = 'selection') {
        return new Promise((resolve, reject) => {
            Word.run(async (context) => {
                try {
                    let content = '';
                    
                    if (source === 'selection') {
                        const selection = context.document.getSelection();
                        selection.load('text');
                        await context.sync();
                        content = selection.text;
                    } else {
                        const body = context.document.body;
                        body.load('text');
                        await context.sync();
                        content = body.text;
                    }
                    
                    resolve(content.trim());
                } catch (error) {
                    reject(error);
                }
            });
        });
    },
    
    // 显示结果
    displayResult: function(text) {
        window.currentResult = text;
        const resultBox = document.getElementById('resultBox');
        if (resultBox) {
            resultBox.innerHTML = text;
            resultBox.classList.remove('loading');
            console.log('结果已显示在结果框中');
        } else {
            console.error('未找到结果框元素');
        }
    },
    
    // 测试按钮状态
    testButtonStates: function() {
        const buttons = [
            { id: 'startBtn', name: '开始处理' },
            { id: 'continueBtn', name: '继续对话' },
            { id: 'insertBtn', name: '插入' },
            { id: 'copyBtn', name: '复制' },
            { id: 'clearBtn', name: '清空' }
        ];
        
        console.log('按钮状态检查:');
        buttons.forEach(({ id, name }) => {
            const btn = document.getElementById(id);
            if (btn) {
                console.log(`✅ ${name}按钮: 存在, 禁用状态: ${btn.disabled}`);
            } else {
                console.log(`❌ ${name}按钮: 不存在`);
            }
        });
    },
    
    // 测试API调用
    testApiCall: async function() {
        console.log('🌐 开始测试API调用...');
        
        try {
            // 检查API配置
            if (typeof API_CONFIG === 'undefined') {
                throw new Error('API配置未加载');
            }
            
            console.log('API配置信息:');
            console.log('- Base URL:', API_CONFIG.baseUrl);
            console.log('- User ID:', API_CONFIG.userId);
            console.log('- Timeout:', API_CONFIG.timeout);
            
            // 测试创建对话
            console.log('测试创建对话...');
            const createUrl = `${API_CONFIG.baseUrl}/v1/conversation`;
            const createData = {
                user_id: API_CONFIG.userId
            };
            
            const createResponse = await fetch(createUrl, {
                method: 'POST',
                headers: API_CONFIG.headers,
                body: JSON.stringify(createData),
                signal: AbortSignal.timeout(API_CONFIG.timeout)
            });
            
            if (!createResponse.ok) {
                throw new Error(`创建对话失败: ${createResponse.status}`);
            }
            
            const createResult = await createResponse.json();
            console.log('创建对话响应:', createResult);
            
            // 测试发送消息
            if (createResult.conversation_id) {
                console.log('测试发送消息...');
                const chatUrl = `${API_CONFIG.baseUrl}/v2/conversation/message`;
                const chatData = {
                    user_id: API_CONFIG.userId,
                    conversation_id: createResult.conversation_id,
                    message: "Hello, this is a test message."
                };
                
                const chatResponse = await fetch(chatUrl, {
                    method: 'POST',
                    headers: API_CONFIG.headers,
                    body: JSON.stringify(chatData),
                    signal: AbortSignal.timeout(API_CONFIG.timeout)
                });
                
                if (!chatResponse.ok) {
                    throw new Error(`发送消息失败: ${chatResponse.status}`);
                }
                
                const chatResult = await chatResponse.json();
                console.log('发送消息响应:', chatResult);
                
                console.log('🎉 API调用测试成功！');
                return { success: true, result: chatResult };
            }
            
        } catch (error) {
            console.error('❌ API调用测试失败:', error);
            return { success: false, error: error.message };
        }
    },
    
    // 完整流程测试
    fullFlowTest: async function() {
        console.log('🚀 开始完整流程测试...');
        
        try {
            // 1. 基础功能测试
            const basicTest = await this.testBasicFlow();
            if (!basicTest.success) {
                throw new Error(`基础功能测试失败: ${basicTest.error}`);
            }
            
            // 2. API调用测试
            const apiTest = await this.testApiCall();
            if (!apiTest.success) {
                console.warn('⚠️ API调用测试失败，但继续其他测试');
                console.warn('错误信息:', apiTest.error);
            }
            
            // 3. 插入功能测试
            console.log('3️⃣ 测试插入功能...');
            await this.testInsertFunction();
            
            console.log('🎉 完整流程测试完成！');
            
        } catch (error) {
            console.error('❌ 完整流程测试失败:', error);
        }
    },
    
    // 测试插入功能
    testInsertFunction: async function() {
        if (!window.currentResult) {
            console.log('⚠️ 没有结果可插入，跳过插入测试');
            return;
        }
        
        try {
            await new Promise((resolve, reject) => {
                Word.run(async (context) => {
                    try {
                        const selection = context.document.getSelection();
                        const testText = `[测试插入] ${new Date().toLocaleString()}`;
                        selection.insertText(testText, Word.InsertLocation.after);
                        await context.sync();
                        resolve();
                    } catch (error) {
                        reject(error);
                    }
                });
            });
            
            console.log('✅ 插入功能测试成功');
            
        } catch (error) {
            console.error('❌ 插入功能测试失败:', error);
        }
    }
};

// 快捷方式
window.debugTest = window.debugTools.testBasicFlow;
window.debugApi = window.debugTools.testApiCall;
window.debugFull = window.debugTools.fullFlowTest;

console.log('🔧 调试工具已加载! 可用命令:');
console.log('- debugTest(): 测试基础功能');
console.log('- debugApi(): 测试API调用');
console.log('- debugFull(): 完整流程测试');
console.log('- debugTools.testWordConnection(): 测试Word连接'); 