/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// 应用状态
let currentTool = 'custom';
let currentContentSource = 'selection';
let currentInsertPosition = 'replace'; // 当前选中的插入位置
let currentResult = '';
let conversationHistory = [];
let currentConversationId = null; // GPTBots对话ID
let isInitialized = false; // 防止重复初始化

// 引入API配置
// 注意：在HTML文件中需要先引入 api-config.js

// Predefined AI tool prompts
const AI_TOOLS = {
    translate: {
        name: '翻译',
        prompt: 'NO.001\n\n{content}'
    },
    polish: {
        name: '润色',
        prompt: 'NO.002：\n\n{content}'
    },
    academic: {
        name: '审批建议',
        prompt: 'NO.003：\n\n{content}'
    },
    summary: {
        name: '总结',
        prompt: 'NO.004：\n\n{content}'
    },

    custom: {
        name: '自定义需求',
        prompt: '{userInput}\n\n内容：\n{content}'
    }
};

// 初始化应用
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
        // 确保DOM完全加载后再初始化
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initializeApp);
        } else {
            initializeApp();
        }
    }
});

function initializeApp() {
    // 防止重复初始化
    if (isInitialized) {
        console.log('⚠️ 应用已初始化，忽略重复初始化');
        return;
    }
    
    console.log('开始初始化 GPTBots Copilot ...');
    
    try {
        // 检查API配置是否已加载
        if (typeof API_CONFIG === 'undefined') {
            throw new Error('API配置文件未正确加载');
        }
        
        // 检查必要的DOM元素是否存在
        const requiredElements = [
            'insertBtn', 'copyBtn',
            'resultBox', 'errorMessage', 'successMessage'
        ];
        
        for (const elementId of requiredElements) {
            if (!document.getElementById(elementId)) {
                throw new Error(`必需的DOM元素未找到: ${elementId}`);
            }
        }
        
        // 检查AI工具按钮
        const aiToolBtns = document.querySelectorAll('.ai-tool-btn');
        console.log(`发现 ${aiToolBtns.length} 个AI工具按钮`);
        
        // 检查内容源按钮
        const contentSourceBtns = document.querySelectorAll('.content-source-btn');
        console.log(`发现 ${contentSourceBtns.length} 个内容源按钮`);
        
        // 绑定事件监听器
        bindEventListeners();
        
        // 初始化UI状态
        updateUI();
        
        // 显示API配置信息
        console.log('GPTBots Copilot 已初始化');
        console.log('API配置:', {
            baseUrl: API_CONFIG.baseUrl,
            createConversationUrl: getCreateConversationUrl(),
            chatUrl: getChatUrl(),
            userId: API_CONFIG.userId
        });
        
        showSuccessMessage('🎉 GPTBots Copilot就绪！');
        
        // 更新结果框显示
        const resultBox = document.getElementById('resultBox');
        if (resultBox) {
            const resultContent = document.getElementById('resultContent');
            if (resultContent) {
                resultContent.textContent = '点击 "开始处理" ';
            } else {
                resultBox.textContent = '点击 "开始处理" ';
            }
            resultBox.classList.remove('loading');
        }
        
        // 初始化自定义输入框显示状态（默认选中custom）
        if (currentTool === 'custom') {
            showCustomInput();
        } else {
            hideCustomInput();
        }
        
        // 初始化按钮状态
        const insertBtn = document.getElementById('insertBtn');
        if (insertBtn) {
            insertBtn.disabled = true; // 初始禁用插入按钮
        }
        
        console.log('GPTBots Copilot 初始化完成！');
        
        // 标记为已初始化
        isInitialized = true;
        
    } catch (error) {
        console.error('初始化失败:', error);
        
        // 在控制台显示详细的调试信息，不在用户界面显示技术错误
        console.log('调试信息:');
        console.log('- API_CONFIG 是否存在:', typeof API_CONFIG !== 'undefined');
        console.log('- 当前DOM状态:', document.readyState);
        console.log('- AI工具按钮数量:', document.querySelectorAll('.ai-tool-btn').length);
        console.log('- 内容源按钮数量:', document.querySelectorAll('.content-source-btn').length);
        console.log('- 错误详情:', error.message);
        
        // 显示友好的初始化状态给用户
        const resultBox = document.getElementById('resultBox');
        if (resultBox) {
            resultBox.innerHTML = `
                <div style="text-align: center; color: #f59e0b; font-weight: 500;">
                    ⚡ GPTBots Copilot初始化中...
                </div>
            `;
        }
        
        // 显示友好的提示而不是技术错误
        showUserFriendlyMessage('GPTBots Copilot初始化中，请稍后...');
    }
}

function bindEventListeners() {
    console.log('开始绑定事件监听器...');
    
    // AI工具按钮
    const aiToolBtns = document.querySelectorAll('.ai-tool-btn');
    console.log(`绑定 ${aiToolBtns.length} 个AI工具按钮:`);
    aiToolBtns.forEach((btn, index) => {
        const toolName = btn.getAttribute('data-tool');
        console.log(`  - 按钮 ${index + 1}: ${btn.textContent} (data-tool: ${toolName})`);
        
        // 清除可能存在的旧事件监听器
        const newBtn = btn.cloneNode(true);
        btn.parentNode.replaceChild(newBtn, btn);
        
        newBtn.addEventListener('click', (event) => {
            console.log(`AI工具按钮被点击: ${newBtn.textContent} (${toolName})`);
            handleToolSelection(event);
        });
    });
    
    // 内容源选择按钮
    const contentSourceBtns = document.querySelectorAll('.content-source-btn');
    console.log(`绑定 ${contentSourceBtns.length} 个内容源按钮:`);
    contentSourceBtns.forEach((btn, index) => {
        const sourceName = btn.getAttribute('data-source');
        console.log(`  - 按钮 ${index + 1}: ${btn.textContent} (data-source: ${sourceName})`);
        
        // 清除可能存在的旧事件监听器
        const newBtn = btn.cloneNode(true);
        btn.parentNode.replaceChild(newBtn, btn);
        
        newBtn.addEventListener('click', (event) => {
            console.log(`内容源按钮被点击: ${newBtn.textContent} (${sourceName})`);
            handleContentSourceSelection(event);
        });
    });
    
    // 主要操作按钮（已移除不存在的按钮）
    console.log('跳过不存在的主要操作按钮绑定');
    
    // 结果操作按钮
    console.log('绑定结果操作按钮:');
    const insertBtn = document.getElementById('insertBtn');
    if (insertBtn) {
        // 清除可能存在的旧事件监听器
        insertBtn.replaceWith(insertBtn.cloneNode(true));
        const newInsertBtn = document.getElementById('insertBtn');
        newInsertBtn.addEventListener('click', () => {
            console.log('插入文档按钮被点击');
            handleInsert();
        });
        console.log('  - 插入文档按钮已绑定');
    }
    
    const copyBtn = document.getElementById('copyBtn');
    if (copyBtn) {
        // 清除可能存在的旧事件监听器
        copyBtn.replaceWith(copyBtn.cloneNode(true));
        const newCopyBtn = document.getElementById('copyBtn');
        newCopyBtn.addEventListener('click', () => {
            console.log('开始处理按钮被点击');
            handleStart();
        });
        console.log('  - 开始处理按钮已绑定（使用copyBtn）');
    }
    
    // 插入位置按钮
    const insertPositionBtns = document.querySelectorAll('.insert-position-btn');
    console.log(`绑定 ${insertPositionBtns.length} 个插入位置按钮:`);
    insertPositionBtns.forEach((btn, index) => {
        const position = btn.getAttribute('data-position');
        console.log(`  - 按钮 ${index + 1}: ${btn.textContent} (data-position: ${position})`);
        
        // 清除可能存在的旧事件监听器
        const newBtn = btn.cloneNode(true);
        btn.parentNode.replaceChild(newBtn, btn);
        
        newBtn.addEventListener('click', (event) => {
            console.log(`插入位置按钮被点击: ${newBtn.textContent} (${position})`);
            handleInsertPositionSelection(event);
        });
    });
    
    // clearBtn 已移除（HTML中不存在）
    console.log('  - 清空按钮不存在，已跳过绑定');
    
    console.log('事件监听器绑定完成！');
}

function handleToolSelection(event) {
    console.log('handleToolSelection 被调用');
    console.log('点击的元素:', event.target);
    console.log('元素内容:', event.target.textContent);
    
    try {
        // 更新选中状态
        document.querySelectorAll('.ai-tool-btn').forEach(btn => {
            btn.classList.remove('selected');
        });
        event.target.classList.add('selected');
        
        // 更新当前工具
        const newTool = event.target.getAttribute('data-tool');
        console.log('选择的工具:', newTool);
        console.log('之前的工具:', currentTool);
        
        currentTool = newTool;
        
        // 如果是自定义工具，显示输入框
        if (currentTool === 'custom') {
            showCustomInput();
            console.log('显示自定义需求输入框');
        } else {
            hideCustomInput();
            console.log('隐藏自定义需求输入框');
        }
        
        
        // 更新UI状态
        updateUI();
        
        console.log(`工具选择完成: ${currentTool}`);
        
    } catch (error) {
        console.error('处理工具选择时出错:', error);
        showUserFriendlyMessage('Tool selection failed, please try again');
    }
}

function handleContentSourceSelection(event) {
    console.log('handleContentSourceSelection 被调用');
    console.log('点击的元素:', event.target);
    console.log('元素内容:', event.target.textContent);
    
    try {
        // 更新选中状态
        document.querySelectorAll('.content-source-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        event.target.classList.add('active');
        
        // 更新当前内容源
        const newSource = event.target.getAttribute('data-source');
        console.log('选择的内容源:', newSource);
        console.log('之前的内容源:', currentContentSource);
        
        currentContentSource = newSource;
        
        // 更新UI状态
        updateUI();
        
        console.log(`内容源选择完成: ${currentContentSource}`);
        
    } catch (error) {
        console.error('处理内容源选择时出错:', error);
        showUserFriendlyMessage('Content source selection failed, please try again');
    }
}

function handleInsertPositionSelection(event) {
    console.log('handleInsertPositionSelection 被调用');
    console.log('点击的元素:', event.target);
    console.log('元素内容:', event.target.textContent);
    
    try {
        // 更新选中状态
        document.querySelectorAll('.insert-position-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        event.target.classList.add('active');
        
        // 更新当前插入位置
        const newPosition = event.target.getAttribute('data-position');
        console.log('选择的插入位置:', newPosition);
        console.log('之前的插入位置:', currentInsertPosition);
        
        currentInsertPosition = newPosition;
        
        console.log(`插入位置选择完成: ${currentInsertPosition}`);
        
    } catch (error) {
        console.error('处理插入位置选择时出错:', error);
        showUserFriendlyMessage('Insert position selection failed, please try again');
    }
}

// 开始处理功能（现在使用copyBtn按钮）
async function handleStart() {
    console.log('🚀 开始处理按钮被点击！');
    console.log('当前工具:', currentTool);
    console.log('当前内容源:', currentContentSource);
    
    const startBtn = document.getElementById('copyBtn');
    
    // 防止重复执行 - 如果按钮已禁用说明正在处理中
    if (startBtn && startBtn.disabled) {
        console.log('⚠️ 处理中，忽略重复点击');
        return;
    }
    
    try {
        // 禁用按钮并显示加载状态
        if (startBtn) {
            startBtn.disabled = true;
            startBtn.classList.add('loading');
            startBtn.innerHTML = '<span>⏳</span><span>处理中...</span>';
        }
        
        // 清除之前的消息
        clearMessages();
        
        // 第一步：显示开始状态
        showLoading('📋 正在获取Word内容...');
        
        // 第二步：获取Word内容
        console.log('📋 正在获取Word内容...');
        const content = await getWordContent();
        console.log('📋 获取到的内容:', content);
        console.log('📋 内容长度:', content.length);
        
        if (!content || content.length === 0) {
            throw new Error(`未找到内容。请先${currentContentSource === 'selection' ? '选择一些文本' : '在文档中添加内容'}。`);
        }
        
        // 在控制台显示技术信息
        console.log(`📊 成功获取${currentContentSource === 'selection' ? '选中文本' : '文档内容'}: ${content.length} 个字符`);
        
        // 第三步：获取用户输入
        const userInput = getUserInput();
        console.log('📋 用户输入:', userInput);
        
        // 如果是自定义工具但没有输入，提示用户
        if (currentTool === 'custom' && !userInput) {
            throw new Error('请在输入框中描述你的需求');
        }
        
        // 第四步：构建提示词
        const prompt = buildPrompt(content, userInput);
        console.log('📋 构建的提示词:', prompt);
        
        showLoading('🤖 AI正在处理中...');
        
        // 第五步：调用API
        console.log('📋 开始调用API...');
        const response = await callConversationAPI(prompt, true); // true表示新对话
        console.log('📋 API响应:', response);
        
        if (!response || response.length === 0) {
            throw new Error('AI返回了空响应');
        }
        
        showLoading('✨ 正在准备结果...');
        
        // 第六步：显示结果
        console.log('📊 开始显示AI响应结果...');
        try {
            displayResult(response);
            console.log(`📊 AI处理完成，生成结果: ${response.length} 个字符`);
        } catch (displayError) {
            console.error('❌ 显示结果时出错:', displayError);
            // 即使显示失败，也要保存结果
            currentResult = response;
        }
        
        // 向用户显示友好信息
        try {
            showSuccessMessage(`处理完成！点击 "插入文档" 将结果添加到Word中。`);
        } catch (msgError) {
            console.error('❌ 显示成功消息时出错:', msgError);
        }
        
        // 启用插入按钮
        try {
            const insertBtn = document.getElementById('insertBtn');
            if (insertBtn) {
                insertBtn.disabled = false;
                console.log('✅ 插入按钮已启用');
            }
        } catch (btnError) {
            console.error('❌ 启用插入按钮时出错:', btnError);
        }
        
        console.log('🎉 处理完成！');
        
    } catch (error) {
        console.error('❌ 处理失败:', error);
        
        // 显示详细的调试信息到控制台
        console.log('调试信息:');
        console.log('- 当前工具:', currentTool);
        console.log('- 当前内容源:', currentContentSource);
        console.log('- API配置存在:', typeof API_CONFIG !== 'undefined');
        console.log('- 错误详情:', error.message);
        console.log('- 错误堆栈:', error.stack);
        
        // 显示友好的错误提示
        showUserFriendlyMessage(error.message);
        
        // 显示默认结果框内容
        const resultBox = document.getElementById('resultBox');
        if (resultBox) {
            const resultContent = document.getElementById('resultContent');
            if (resultContent) {
                resultContent.textContent = '处理失败，请检查输入内容后重试';
            }
        }
        
    } finally {
        // 恢复按钮状态
        if (startBtn) {
            startBtn.disabled = false;
            startBtn.classList.remove('loading');
            startBtn.innerHTML = '<span>🚀</span><span>开始处理</span>';
        }
        hideLoading();
    }
}

// handleContinue函数已移除（continueBtn不存在）
async function handleContinue_REMOVED() {
    try {
        // conversationInput不存在，显示提示
        showUserFriendlyMessage('Continue conversation feature requires input field (not implemented)');
        return;
        
    } catch (error) {
        console.error('继续对话失败:', error);
        showUserFriendlyMessage('Chat feature is being prepared, please try again later');
    } finally {
        hideLoading();
    }
}

async function getWordContent() {
    console.log('📋 getWordContent: 开始获取Word内容...');
    console.log('📋 内容源:', currentContentSource);
    
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                let content = '';
                
                if (currentContentSource === 'selection') {
                    console.log('📋 正在获取选中文本...');
                    // 获取选中的文本
                    const selection = context.document.getSelection();
                    selection.load('text');
                    await context.sync();
                    content = selection.text;
                    console.log('📋 选中文本内容:', content);
                    console.log('📋 选中文本长度:', content.length);
                    
                    if (!content || content.trim().length === 0) {
                        throw new Error('No text selected. Please select some text in Word first.');
                    }
                } else {
                    console.log('📋 正在获取整个文档文本...');
                    // 获取整个文档的文本
                    const body = context.document.body;
                    body.load('text');
    await context.sync();
                    content = body.text;
                    console.log('📋 文档内容长度:', content.length);
                    
                    if (!content || content.trim().length === 0) {
                        throw new Error('Document is empty. Please add some content to the document first.');
                    }
                }
                
                const trimmedContent = content.trim();
                console.log('📋 最终内容长度:', trimmedContent.length);
                console.log('📋 内容前100个字符:', trimmedContent.substring(0, 100));
                
                resolve(trimmedContent);
            } catch (error) {
                console.error('📋 获取Word内容失败:', error);
                reject(error);
            }
        });
    });
}

function buildPrompt(content, userInput) {
    const tool = AI_TOOLS[currentTool];
    
    let prompt = tool.prompt;
    
    // 替换模板变量
    prompt = prompt.replace('{content}', content);
    prompt = prompt.replace('{userInput}', userInput || '');
    
    // 使用默认语言（中文）替换语言占位符
    prompt = prompt.replace('{language}', '中文');
    
    return prompt;
}

function getLanguageName(code) {
    const languageMap = {
        'zh': '中文',
        'en': '英文',
        'ja': '日文',
        'ko': '韩文',
        'fr': '法文',
        'de': '德文',
        'es': '西班牙文',
        'ru': '俄文'
    };
    return languageMap[code] || '中文';
}

async function callConversationAPI(prompt, isNewConversation = true) {
    try {
        // 尝试使用本地代理API
        if (typeof window.localProxyAPI !== 'undefined') {
            console.log('🔄 使用本地代理API...');
            
            let conversationId = currentConversationId;
            
            if (isNewConversation || !conversationId) {
                console.log('📞 创建新对话...');
                const createResult = await window.localProxyAPI.createConversation();
                if (createResult.success) {
                    conversationId = createResult.conversationId;
                    currentConversationId = conversationId;
                    console.log('✅ 对话创建成功:', conversationId);
                } else {
                    throw new Error('本地代理创建对话失败');
                }
            }
            
            console.log('📞 发送消息...');
            const messageResult = await window.localProxyAPI.sendMessage(conversationId, prompt);
            if (messageResult.success) {
                console.log('✅ 消息发送成功');
                return messageResult.message;
            } else {
                throw new Error('本地代理发送消息失败');
            }
        }
        
        // 如果本地代理不可用，尝试直接API调用
        // 如果是新对话，需要先创建对话
        if (isNewConversation) {
            conversationHistory = [];
            currentConversationId = null;
            
            // 第一步：创建对话
            console.log('创建新对话...');
            const createResponse = await fetch(getCreateConversationUrl(), {
                method: 'POST',
                headers: API_CONFIG.headers,
                body: JSON.stringify(buildCreateConversationData()),
                signal: AbortSignal.timeout(API_CONFIG.timeout)
            });
            
            if (!createResponse.ok) {
                throw new Error(`创建对话失败: ${createResponse.status} ${createResponse.statusText}`);
            }
            
            const createResult = await createResponse.json();
            console.log('创建对话响应:', createResult);
            
            const parsedCreateResult = parseCreateConversationResponse(createResult);
            
            if (!parsedCreateResult.success) {
                throw new Error(parsedCreateResult.error || '创建对话失败');
            }
            
            currentConversationId = parsedCreateResult.conversationId;
            console.log('对话ID:', currentConversationId);
        }
        
        // 确保有对话ID
        if (!currentConversationId) {
            throw new Error('缺少对话ID，请重新开始对话');
        }
        
        // 添加用户消息到历史记录
        conversationHistory.push({
            role: 'user',
            content: prompt
        });
        
        // 第二步：发送消息
        console.log('发送消息...');
        const chatRequestData = buildChatRequestData(currentConversationId, conversationHistory);
        console.log('消息请求数据:', chatRequestData);
        
        const chatResponse = await fetch(getChatUrl(), {
            method: 'POST',
            headers: API_CONFIG.headers,
            body: JSON.stringify(chatRequestData),
            signal: AbortSignal.timeout(API_CONFIG.timeout)
        });
        
        if (!chatResponse.ok) {
            throw new Error(`发送消息失败: ${chatResponse.status} ${chatResponse.statusText}`);
        }
        
        const chatResult = await chatResponse.json();
        console.log('消息响应:', chatResult);
        
        // 解析消息响应
        const parsedChatResult = parseChatResponse(chatResult);
        
        if (!parsedChatResult.success) {
            throw new Error(parsedChatResult.error || '消息处理失败');
        }
        
        // 添加助手消息到历史记录
        conversationHistory.push({
            role: 'assistant',
            content: parsedChatResult.message
        });
        
        return parsedChatResult.message;
        
    } catch (error) {
        console.error('API调用错误:', error);
        console.log('💡 建议：确保本地代理服务器运行: node local-server.js');
        
        // 抛出错误让上层函数处理
        throw new Error(`API调用失败: ${error.message}`);
    }
}

async function handleInsert() {
    console.log('📝 插入按钮被点击');
    console.log('📝 当前结果长度:', currentResult ? currentResult.length : 0);
    
    if (!currentResult) {
        showUserFriendlyMessage('没有内容可插入，请先点击"开始处理"');
        return;
    }
    
    const insertBtn = document.getElementById('insertBtn');
    
    // 防止重复执行 - 如果按钮已禁用说明正在插入中
    if (insertBtn && insertBtn.disabled) {
        console.log('⚠️ 插入中，忽略重复点击');
        return;
    }
    
    try {
        // 禁用按钮并显示加载状态
        if (insertBtn) {
            insertBtn.disabled = true;
            insertBtn.classList.add('loading');
            insertBtn.innerHTML = '<span>⏳</span><span>插入中...</span>';
        }
        
        let insertType = currentInsertPosition;
        
        // 如果是审批建议功能，强制使用批注模式
        if (currentTool === 'academic') {
            insertType = 'comment';
            console.log('📝 审批建议功能：强制使用批注模式');
        }
        
        console.log('📝 插入类型:', insertType);
        
        showLoading('📝 正在将内容插入Word文档...');
        
        await insertToWordWithType(currentResult, insertType);
        
        const insertTypeText = {
            'replace': '替换选中文本',
            'append': '添加到文档末尾',
            'cursor': '在光标位置插入',
            'comment': '生成批注'
        }[insertType] || '插入';
        
        showSuccessMessage(`内容已成功${insertTypeText}！`);
        console.log('�� 插入成功！');
        
        // 强制清除加载状态
        hideLoading();
        
    } catch (error) {
        console.error('📝 插入失败:', error);
        showUserFriendlyMessage(`插入失败：${error.message}`);
    } finally {
        // 恢复按钮状态
        if (insertBtn) {
            insertBtn.disabled = false;
            insertBtn.classList.remove('loading');
            insertBtn.innerHTML = '<span>📝</span><span>插入文档</span>';
        }
        hideLoading();
    }
}

async function insertToWordWithType(text, insertType) {
    console.log('📝 insertToWordWithType: 开始插入文本');
    console.log('📝 要插入的文本长度:', text.length);
    console.log('📝 插入类型:', insertType);
    
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                
                switch (insertType) {
                    case 'replace':
                        console.log('📝 执行替换选中文本操作');
                        // 替换选中的文本
                        const selection = context.document.getSelection();
                        selection.insertText(text, Word.InsertLocation.replace);
                        break;
                        
                    case 'append':
                        console.log('📝 执行追加到文档末尾操作');
                        // 追加到文档末尾
                        const body = context.document.body;
                        body.insertParagraph('\n' + text, Word.InsertLocation.end);
                        break;
                        
                    case 'cursor':
                        console.log('📝 执行在光标位置插入操作');
                        // 在光标位置插入
                        const range = context.document.getSelection();
                        range.insertText(text, Word.InsertLocation.after);
                        break;
                        
                    case 'comment':
                        console.log('📝 执行生成批注操作');
                        // 为选中文本添加批注
                        const selectionForComment = context.document.getSelection();
                        selectionForComment.load('isEmpty');
                        await context.sync();
                        
                        if (selectionForComment.isEmpty) {
                            console.log('📝 没有选中文本，将在文档末尾插入批注内容');
                            // 如果没有选中文本，在文档末尾插入内容
                            const body = context.document.body;
                            body.insertParagraph('\n【审批建议】\n' + text, Word.InsertLocation.end);
                        } else {
                            console.log('📝 为选中文本添加批注');
                            // 添加批注
                            selectionForComment.insertComment(text);
                        }
                        break;
                        
                    default:
                        throw new Error(`未知的插入类型: ${insertType}`);
                }
                
                console.log('📝 正在同步到Word...');
    await context.sync();
                console.log('📝 插入完成！');
                
                resolve();
            } catch (error) {
                console.error('📝 插入到Word时出错:', error);
                reject(error);
            }
        });
    });
}

// handleCopy函数已移除（copyBtn现在用于开始处理）
function handleCopy_REMOVED() {
    if (!currentResult) {
        showUserFriendlyMessage('No content to copy');
        return;
    }
    
    // 使用现代浏览器的剪贴板API
    if (navigator.clipboard) {
        navigator.clipboard.writeText(currentResult).then(() => {
            showSuccessMessage('Content copied to clipboard');
        }).catch(() => {
            // 降级到传统方法
            fallbackCopy(currentResult);
        });
    } else {
        fallbackCopy(currentResult);
    }
}

function fallbackCopy(text) {
    // 降级复制方法
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.opacity = '0';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            showSuccessMessage('Content copied to clipboard');
        } else {
            showUserFriendlyMessage('Copy function temporarily unavailable, please manually select and copy content from result area');
        }
    } catch (err) {
        showUserFriendlyMessage('Copy function temporarily unavailable, please manually select and copy content from result area');
    }
    
    document.body.removeChild(textArea);
}

function handleClear() {
    console.log('🗑️ 开始清空操作...');
    
    // 分步骤执行，每一步都有独立的错误处理
    
    // 步骤1：清空变量
    try {
        currentResult = '';
        conversationHistory = [];
        currentConversationId = null;
        console.log('✅ 步骤1：变量清空完成');
    } catch (error) {
        console.warn('步骤1失败:', error);
    }
    
    // 步骤2：清空结果框
    try {
        const resultBox = document.getElementById('resultBox');
        if (resultBox) {
            const resultContent = document.getElementById('resultContent');
            if (resultContent) {
                resultContent.textContent = '选择AI工具后点击 "运行" 获取AI响应';
            } else {
                resultBox.textContent = '选择AI工具后点击 "运行" 获取AI响应';
            }
            resultBox.classList.remove('loading');
        }
        console.log('✅ 步骤2：结果框清空完成');
    } catch (error) {
        console.warn('步骤2失败:', error);
    }
    
    // 步骤3：清空输入框
    try {
        const customTextarea = document.getElementById('customInputTextarea');
        if (customTextarea) {
            customTextarea.value = '';
        }
        console.log('✅ 步骤3：自定义输入框清空完成');
    } catch (error) {
        console.warn('步骤3失败:', error);
    }
    
    // 步骤4：清空消息
    try {
        const errorElement = document.getElementById('errorMessage');
        if (errorElement) {
            errorElement.classList.add('hidden');
        }
        
        const successElement = document.getElementById('successMessage');
        if (successElement) {
            successElement.classList.add('hidden');
        }
        console.log('✅ 步骤4：消息清空完成');
    } catch (error) {
        console.warn('步骤4失败:', error);
    }
    
    // 步骤5：显示成功消息（延迟执行）
    setTimeout(() => {
        try {
            const successElement = document.getElementById('successMessage');
            if (successElement) {
                successElement.textContent = 'Results and conversation cleared';
                successElement.classList.remove('hidden');
                
                // 3秒后隐藏
                setTimeout(() => {
                    try {
                        if (successElement) {
                            successElement.classList.add('hidden');
                        }
                    } catch (e) {
                        console.warn('隐藏成功消息失败:', e);
                    }
                }, 3000);
            }
            console.log('✅ 步骤5：成功消息显示完成');
        } catch (error) {
            console.warn('步骤5失败:', error);
        }
    }, 100);
    
    console.log('🎉 清空操作全部完成');
}

function displayResult(result) {
    try {
        console.log('📊 开始显示结果，长度:', result ? result.length : 0);
        
        currentResult = result;
        const resultBox = document.getElementById('resultBox');
        
        if (!resultBox) {
            console.error('❌ 未找到resultBox元素');
            return;
        }
        
        // 清除加载状态
        resultBox.classList.remove('loading');
        
        // 确保结果框有正确的结构
        let resultContent = document.getElementById('resultContent');
        if (!resultContent) {
            resultBox.innerHTML = '<div id="resultContent"></div>';
            resultContent = document.getElementById('resultContent');
        }
        
        if (resultContent) {
            resultContent.textContent = result;
            console.log('✅ 结果已显示在resultContent中');
        } else {
            // 降级处理
            resultBox.innerHTML = `<div id="resultContent">${result}</div>`;
            console.log('✅ 结果已显示在resultBox中（降级处理）');
        }
        
        // 启用插入按钮
        const insertBtn = document.getElementById('insertBtn');
        if (insertBtn) {
            insertBtn.disabled = false;
            console.log('✅ 插入按钮已启用');
        }
        
        console.log('📊 结果显示完成');
        
    } catch (error) {
        console.error('❌ 显示结果时出错:', error);
        console.error('错误堆栈:', error.stack);
        
        // 降级处理：直接在控制台显示结果
        console.log('📊 降级处理 - 结果内容:', result);
    }
}

// 帮助函数：创建加载动画HTML
function createLoadingHTML(message) {
    return `
        <div class="loading-animation">
            <div class="loading-spinner"></div>
            <div class="loading-dots">
                <div class="loading-dot"></div>
                <div class="loading-dot"></div>
                <div class="loading-dot"></div>
            </div>
        </div>
        <div class="loading-text">${message}</div>
    `;
}

function showLoading(message) {
    const resultBox = document.getElementById('resultBox');
    
    // 创建现代化的加载动画
    resultBox.innerHTML = createLoadingHTML(message);
    resultBox.classList.add('loading');
    
    // 禁用按钮（startBtn和continueBtn不存在，跳过）
    console.log('跳过禁用不存在的按钮');
    
    console.log('🔄 显示加载状态:', message);
}

function hideLoading() {
    const resultBox = document.getElementById('resultBox');
    if (resultBox) {
        resultBox.classList.remove('loading');
        
        // 如果结果框仍然显示加载动画，清除它
        if (resultBox.innerHTML.includes('loading-spinner') || resultBox.innerHTML.includes('⏳')) {
            // 如果有当前结果，显示结果；否则显示默认提示
            if (currentResult) {
                displayResult(currentResult);
            } else {
                const resultContent = document.getElementById('resultContent');
                if (resultContent) {
                    resultContent.textContent = '选择AI工具后点击 "开始处理" 获取Agent响应';
                } else {
                    resultBox.innerHTML = '<div id="resultContent">选择AI工具后点击 "开始处理" 获取Agent响应</div>';
                }
            }
        }
    }
    
    // 启用按钮（startBtn和continueBtn不存在，跳过）
    console.log('跳过启用不存在的按钮');
    
    console.log('✅ 隐藏加载状态');
}

function showErrorMessage(message) {
    // 只在控制台显示技术错误信息
    console.warn('❌ 错误信息 (仅控制台显示):', message);
    
    // 不在用户界面显示错误信息
    // 如果需要向用户显示信息，使用 showUserFriendlyMessage
}

function showUserFriendlyMessage(message) {
    // 新增函数：专门用于显示用户友好的信息
    try {
        const successElement = document.getElementById('successMessage');
        if (successElement) {
            successElement.textContent = message;
            successElement.classList.remove('hidden');
            
            // 5秒后自动隐藏
            setTimeout(() => {
                if (successElement) {
                    successElement.classList.add('hidden');
                }
            }, 5000);
        }
        
        console.log('💬 用户提示:', message);
    } catch (error) {
        console.warn('显示用户友好消息时出错:', error);
        console.log('💬 用户提示:', message);
    }
}

function showSuccessMessage(message) {
    try {
        const successElement = document.getElementById('successMessage');
        if (successElement) {
            successElement.textContent = message;
            successElement.classList.remove('hidden');
            
            // 3秒后自动隐藏
            setTimeout(() => {
                if (successElement) {
                    successElement.classList.add('hidden');
                }
            }, 3000);
        }
        
        console.log('✅ 成功消息:', message);
    } catch (error) {
        console.warn('显示成功消息时出错:', error);
        console.log('✅ 成功消息:', message);
    }
}

function clearMessages() {
    try {
        const errorElement = document.getElementById('errorMessage');
        if (errorElement) {
            errorElement.classList.add('hidden');
        }
        
        const successElement = document.getElementById('successMessage');
        if (successElement) {
            successElement.classList.add('hidden');
        }
    } catch (error) {
        console.warn('清除消息时出错:', error);
    }
}

function updateUI() {
    try {
        // 更新自定义输入框显示
        if (currentTool === 'custom') {
            showCustomInput();
        } else {
            hideCustomInput();
        }
        
        console.log('UI状态已更新');
    } catch (error) {
        console.warn('更新UI时出错:', error);
    }
}

// 显示自定义需求输入框
function showCustomInput() {
    const container = document.getElementById('customInputContainer');
    if (container) {
        container.classList.remove('hidden');
        
        // 聚焦到输入框
        const textarea = document.getElementById('customInputTextarea');
        if (textarea) {
            setTimeout(() => {
                textarea.focus();
            }, 100);
        }
    }
}

// 隐藏自定义需求输入框
function hideCustomInput() {
    const container = document.getElementById('customInputContainer');
    if (container) {
        container.classList.add('hidden');
    }
}

// 获取用户输入
function getUserInput() {
    if (currentTool === 'custom') {
        const textarea = document.getElementById('customInputTextarea');
        if (textarea) {
            return textarea.value.trim();
        }
    }
    return '';
}

// 调试工具函数 - 在浏览器控制台中可以手动调用
window.debugWordGPT = {
    // 测试按钮绑定
    testButtonBindings: function() {
        console.log('=== 测试按钮绑定 ===');
        
        const aiToolBtns = document.querySelectorAll('.ai-tool-btn');
        console.log(`AI工具按钮数量: ${aiToolBtns.length}`);
        aiToolBtns.forEach((btn, i) => {
            console.log(`  ${i+1}. ${btn.textContent} - data-tool: ${btn.getAttribute('data-tool')}`);
        });
        
        const contentBtns = document.querySelectorAll('.content-source-btn');
        console.log(`内容源按钮数量: ${contentBtns.length}`);
        contentBtns.forEach((btn, i) => {
            console.log(`  ${i+1}. ${btn.textContent} - data-source: ${btn.getAttribute('data-source')}`);
        });
        
        const actionBtns = ['copyBtn', 'insertBtn'];
        console.log('操作按钮:');
        actionBtns.forEach(id => {
            const btn = document.getElementById(id);
            const btnName = id === 'copyBtn' ? '开始处理' : '插入文档';
            console.log(`  ${id} (${btnName}): ${btn ? '找到' : '未找到'}`);
        });
    },
    
    // 手动触发工具选择
    selectTool: function(toolName) {
        console.log(`尝试选择工具: ${toolName}`);
        const btn = document.querySelector(`[data-tool="${toolName}"]`);
        if (btn) {
            btn.click();
            console.log('按钮点击成功');
        } else {
            console.log('未找到按钮');
        }
    },
    
    // 手动触发内容源选择
    selectSource: function(sourceName) {
        console.log(`尝试选择内容源: ${sourceName}`);
        const btn = document.querySelector(`[data-source="${sourceName}"]`);
        if (btn) {
            btn.click();
            console.log('按钮点击成功');
        } else {
            console.log('未找到按钮');
        }
    },
    
    // 显示当前状态
    showStatus: function() {
        console.log('=== 当前状态 ===');
        console.log('当前工具:', currentTool);
        console.log('当前内容源:', currentContentSource);
        console.log('对话ID:', currentConversationId);
        console.log('对话历史长度:', conversationHistory.length);
        console.log('当前结果长度:', currentResult.length);
        
        // 显示自定义输入状态
        if (currentTool === 'custom') {
            const userInput = getUserInput();
            console.log('自定义需求输入:', userInput || '(空)');
        }
    },
    
    // 重新初始化
    reinitialize: function() {
        console.log('重新初始化...');
        initializeApp();
    },
    
    // 快速测试整个流程
    quickTest: function() {
        console.log('🧪 开始快速测试...');
        
        // 测试1: 检查是否有选中文本
        Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            
            if (selection.text && selection.text.trim().length > 0) {
                console.log('✅ 发现选中文本:', selection.text);
                console.log('📝 文本长度:', selection.text.length);
                
                // 自动选择翻译工具（startBtn不存在，无法自动处理）
                debugWordGPT.selectTool('translate');
                
                console.log('💡 startBtn不存在，无法自动开始处理');
                
            } else {
                console.log('❌ 没有选中文本');
                console.log('💡 Please select text in Word first, then run debugWordGPT.quickTest() again');
                
                // 显示提示
                const resultBox = document.getElementById('resultBox');
                if (resultBox) {
                    resultBox.textContent = 'Please select text in Word first';
                }
            }
        }).catch(error => {
            console.error('❌ 快速测试失败:', error);
        });
    },
    
    // 测试Word连接
    testWordConnection: function() {
        console.log('🔗 测试Word连接...');
        
        Word.run(async (context) => {
            console.log('✅ Word连接成功');
            
            // 获取选中文本
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            
            console.log('选中文本:', selection.text);
            console.log('选中文本长度:', selection.text.length);
            
            // 获取文档内容
            const body = context.document.body;
            body.load('text');
            await context.sync();
            
            console.log('文档总长度:', body.text.length);
            console.log('文档前100个字符:', body.text.substring(0, 100));
            
            return true;
        }).catch(error => {
            console.error('❌ Word连接失败:', error);
            return false;
        });
    }
};

// 添加全局错误处理器，防止未捕获的错误显示弹窗
window.addEventListener('error', function(event) {
    console.error('🚫 全局错误捕获:', event.error);
    console.error('错误详情:', {
        message: event.message,
        filename: event.filename,
        lineno: event.lineno,
        colno: event.colno,
        error: event.error
    });
    
    // 阻止默认的错误处理（防止弹窗）
    event.preventDefault();
    return true;
});

// 捕获Promise中的未处理错误
window.addEventListener('unhandledrejection', function(event) {
    console.error('🚫 未处理的Promise错误:', event.reason);
    
    // 阻止默认的错误处理（防止弹窗）
    event.preventDefault();
    return true;
});

console.log('调试工具已加载！在控制台输入 debugWordGPT.testButtonBindings() 来测试按钮绑定');
console.log('已启用全局错误捕获，防止弹窗错误');
console.log('✅ 已启用防重复执行保护机制');
