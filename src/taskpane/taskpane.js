/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// 应用状态
let currentTool = null;
let currentContentSource = 'selection';
let currentInsertPosition = 'replace'; // 当前选中的插入位置
let currentResult = '';
let conversationHistory = [];
let currentConversationId = null; // GPTBots对话ID
let isInitialized = false; // 防止重复初始化
let currentLanguage = 'zh-cn'; // 当前选择的语言
let selectedTranslateLanguage = null; // 选择的翻译目标语言


// 引入API配置
// 注意：在HTML文件中需要先引入 api-config.js

// 多语言配置
const LANGUAGE_TEXTS = {
    'zh-cn': {
        languageSettings: '🌐 语言设置',
        skills: 'Agent 技能',
        targetContent: '目标内容',
        resultPreview: '结果预览',
        insertPosition: '生成位置',
        'btn.translate': '深度翻译',
        'btn.polish': '内容润色',
        'btn.academic': '审批建议',
        'btn.summary': '总结摘要',
        'btn.grammar': '语法修正',
        'btn.enterprise': '企业数据',
        'btn.selection': '选中文本',
        'btn.document': '整个文档',
        'btn.start': '开始处理',
        'btn.insert': '插入文档',
        'btn.replace': '替换选中文本',
        'btn.append': '添加至末尾',
        'btn.cursor': '光标位置插入',
        'btn.comment': '生成批注',
        'placeholder.custom': 'GPTBots会根据你的需求生成内容...',
        'translate.selectLanguage': '选择翻译目标语言',
        'lang.english': 'English',
        'lang.chinese': '中文',
        'lang.japanese': '日本語',
        'lang.korean': '한국어',
        'lang.thai': 'ไทย',
        'lang.french': 'Français',
        'lang.german': 'Deutsch',
        'lang.spanish': 'Español',
        'lang.traditional': '繁體中文',
        'translate.to': '翻译成',
        'apiKey.title': 'GPTBots API Key',
        'apiKey.placeholder': 'app-xxxxxxxxxxxxxxxxxxxxxxx',

        'enterprise.inputLabel': '请描述您需要的企业数据',
        'enterprise.inputPlaceholder': '例如：上季度销售数据、去年财务报表、客户满意度调研等...',
        'custom.inputLabel': '请描述您的具体需求',
        'custom.inputPlaceholder': '例如：总结要点、修正语法错误、特定格式要求等...',
        'btn.back': '返回'
    },
    'en': {
        languageSettings: '🌐 Language Settings',
        skills: 'Agent Skills',
        targetContent: 'Target Content',
        resultPreview: 'Result Preview',
        insertPosition: 'Insert Position',
        'btn.translate': 'Deep Translation',
        'btn.polish': 'Content Polish',
        'btn.academic': 'Review Suggestions',
        'btn.summary': 'Summary',
        'btn.grammar': 'Grammar Fix',
        'btn.enterprise': 'Enterprise Data',
        'btn.selection': 'Selected Text',
        'btn.document': 'Entire Document',
        'btn.start': 'Start Processing',
        'btn.insert': 'Insert to Document',
        'btn.replace': 'Replace Selected Text',
        'btn.append': 'Append to End',
        'btn.cursor': 'Insert at Cursor',
        'btn.comment': 'Add Comment',
        'placeholder.custom': 'GPTBots will generate content based on your requirements...',
        'translate.selectLanguage': 'Select Target Language',
        'lang.english': 'English',
        'lang.chinese': '中文',
        'lang.japanese': '日本語',
        'lang.korean': '한국어',
        'lang.thai': 'ไทย',
        'lang.french': 'Français',
        'lang.german': 'Deutsch',
        'lang.spanish': 'Español',
        'lang.traditional': '繁體中文',
        'translate.to': 'Translate to',
        'apiKey.title': 'GPTBots API Key',
        'apiKey.placeholder': 'app-xxxxxxxxxxxxxxxxxxxxxxx',

        'enterprise.inputLabel': 'Please describe the enterprise data you need',
        'enterprise.inputPlaceholder': 'e.g.: Last quarter sales data, annual financial reports, customer satisfaction surveys...',
        'custom.inputLabel': 'Please describe your specific requirements',
        'custom.inputPlaceholder': 'e.g.: Summarize key points, fix grammar errors, specific format requirements...',
        'btn.back': 'Back'
    },
    'th': {
        languageSettings: '🌐 การตั้งค่าภาษา',
        skills: 'ทักษะ Agent',
        targetContent: 'เนื้อหาเป้าหมาย',
        resultPreview: 'ตัวอย่างผลลัพธ์',
        insertPosition: 'ตำแหน่งการแทรก',
        'btn.translate': 'การแปลเชิงลึก',
        'btn.polish': 'ปรับปรุงเนื้อหา',
        'btn.academic': 'ข้อเสนอแนะการอนุมัติ',
        'btn.summary': 'สรุปย่อ',
        'btn.grammar': 'แก้ไขไวยากรณ์',
        'btn.enterprise': 'ข้อมูลองค์กร',
        'btn.selection': 'ข้อความที่เลือก',
        'btn.document': 'เอกสารทั้งหมด',
        'btn.start': 'เริ่มประมวลผล',
        'btn.insert': 'แทรกในเอกสาร',
        'btn.replace': 'แทนที่ข้อความที่เลือก',
        'btn.append': 'เพิ่มที่ท้าย',
        'btn.cursor': 'แทรกที่เคอร์เซอร์',
        'btn.comment': 'เพิ่มความคิดเห็น',
        'placeholder.custom': 'GPTBots จะสร้างเนื้อหาตามความต้องการของคุณ...',
        'translate.selectLanguage': 'เลือกภาษาเป้าหมาย',
        'lang.english': 'English',
        'lang.chinese': '中文',
        'lang.japanese': '日本語',
        'lang.korean': '한국어',
        'lang.thai': 'ไทย',
        'lang.french': 'Français',
        'lang.german': 'Deutsch',
        'lang.spanish': 'Español',
        'lang.traditional': '繁體中文',
        'translate.to': 'แปลเป็น',
        'apiKey.title': 'GPTBots API Key',
        'apiKey.placeholder': 'app-xxxxxxxxxxxxxxxxxxxxxxx',

        'enterprise.inputLabel': 'โปรดอธิบายข้อมูลองค์กรที่คุณต้องการ',
        'enterprise.inputPlaceholder': 'เช่น: ข้อมูลการขายไตรมาสที่แล้ว รายงานการเงินประจำปี การสำรวจความพึงพอใจของลูกค้า...',
        'custom.inputLabel': 'โปรดอธิบายความต้องการเฉพาะของคุณ',
        'custom.inputPlaceholder': 'เช่น: สรุปประเด็นสำคัญ แก้ไขข้อผิดพลาดทางไวยากรณ์ ข้อกำหนดรูปแบบเฉพาะ...',
        'btn.back': 'กลับ'
    },
    'ja': {
        languageSettings: '🌐 言語設定',
        skills: 'Agent スキル',
        targetContent: '対象コンテンツ',
        resultPreview: '結果プレビュー',
        insertPosition: '挿入位置',
        'btn.translate': '深度翻訳',
        'btn.polish': 'コンテンツ校正',
        'btn.academic': '承認提案',
        'btn.summary': '要約',
        'btn.grammar': '文法修正',
        'btn.enterprise': '企業データ',
        'btn.selection': '選択テキスト',
        'btn.document': '文書全体',
        'btn.start': '処理開始',
        'btn.insert': '文書に挿入',
        'btn.replace': '選択テキストを置換',
        'btn.append': '末尾に追加',
        'btn.cursor': 'カーソル位置に挿入',
        'btn.comment': 'コメント追加',
        'placeholder.custom': 'GPTBotsがあなたの要件に基づいてコンテンツを生成します...',
        'translate.selectLanguage': '翻訳先言語を選択',
        'lang.english': 'English',
        'lang.chinese': '中文',
        'lang.japanese': '日本語',
        'lang.korean': '한국어',
        'lang.thai': 'ไทย',
        'lang.french': 'Français',
        'lang.german': 'Deutsch',
        'lang.spanish': 'Español',
        'lang.traditional': '繁體中文',
        'translate.to': '翻訳先：',
        'apiKey.title': 'GPTBots API Key',
        'apiKey.placeholder': 'app-xxxxxxxxxxxxxxxxxxxxxxx',

        'enterprise.inputLabel': '必要な企業データを説明してください',
        'enterprise.inputPlaceholder': '例：前四半期の売上データ、年次財務報告書、顧客満足度調査など...',
        'custom.inputLabel': '具体的な要件を説明してください',
        'custom.inputPlaceholder': '例：要点をまとめる、文法エラーを修正する、特定の形式要件など...',
        'btn.back': '戻る'
    },
    'zh-tw': {
        languageSettings: '🌐 語言設置',
        skills: 'Agent 技能',
        targetContent: '目標內容',
        resultPreview: '結果預覽',
        insertPosition: '生成位置',
        'btn.translate': '深度翻譯',
        'btn.polish': '內容潤色',
        'btn.academic': '審批建議',
        'btn.summary': '總結摘要',
        'btn.grammar': '語法修正',
        'btn.enterprise': '企業數據',
        'btn.selection': '選中文字',
        'btn.document': '整個文件',
        'btn.start': '開始處理',
        'btn.insert': '插入文件',
        'btn.replace': '替換選中文字',
        'btn.append': '添加至末尾',
        'btn.cursor': '游標位置插入',
        'btn.comment': '生成批註',
        'placeholder.custom': 'GPTBots會根據你的需求生成內容...',
        'translate.selectLanguage': '選擇翻譯目標語言',
        'lang.english': 'English',
        'lang.chinese': '中文',
        'lang.japanese': '日本語',
        'lang.korean': '한국어',
        'lang.thai': 'ไทย',
        'lang.french': 'Français',
        'lang.german': 'Deutsch',
        'lang.spanish': 'Español',
        'lang.traditional': '繁體中文',
        'translate.to': '翻譯成',
        'apiKey.title': 'GPTBots API Key',
        'apiKey.placeholder': 'app-xxxxxxxxxxxxxxxxxxxxxxx',

        'enterprise.inputLabel': '請描述您需要的企業數據',
        'enterprise.inputPlaceholder': '例如：上季度銷售數據、去年財務報表、客戶滿意度調研等...',
        'custom.inputLabel': '請描述您的具體需求',
        'custom.inputPlaceholder': '例如：總結要點、修正語法錯誤、特定格式要求等...',
        'btn.back': '返回'
    }
};

// Predefined AI tool prompts
const AI_TOOLS = {
    translate: {
        name: '翻译',
        prompt: 'NO.001: 请翻译为{targetLanguage}：{content}'
    },
    polish: {
        name: '润色',
        prompt: 'NO.002: {content}'
    },
    academic: {
        name: '审批建议',
        prompt: 'NO.003: {content}'
    },
    academicDocument: {
        name: '审批建议（整个文档）',
        prompt: 'NO.010: {content}'
    },
    summary: {
        name: '总结',
        prompt: 'NO.004: {userInput}：内容：{content}',
        needsInput: true
    },
    grammar: {
        name: '修改语法',
        prompt: 'NO.005: {userInput}：内容：{content}',
        needsInput: true
    },
    enterprise: {
        name: '企业数据',
        prompt: 'NO.006: {userInput}：内容：{content}',
        needsInput: true
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
        
        
        // 更新结果框显示
        const resultBox = document.getElementById('resultBox');
        if (resultBox) {
            const resultContent = document.getElementById('resultContent');
            if (resultContent) {
                resultContent.textContent = 'GPTBots Copilot ';
            } else {
                resultBox.textContent = 'GPTBots Copilot';
            }
            resultBox.classList.remove('loading');
        }
        
        // 初始化时隐藏输入框
        hideCustomInput();
        hideEnterpriseInput();
        
        // 初始化按钮状态
        const insertBtn = document.getElementById('insertBtn');
        if (insertBtn) {
            insertBtn.disabled = true; // 初始禁用插入按钮
        }
        
        console.log('GPTBots Copilot 初始化完成！');
        
        // 标记为已初始化
        isInitialized = true;
        
        // 初始化语言显示
        updateLanguageDisplay();
        
        // 初始化API Key掩码
        setTimeout(() => {
            maskApiKey();
        }, 100);
        
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
    
    // 地球图标语言选择器
    const languageGlobeBtn = document.getElementById('languageGlobeBtn');
    const languageDropdown = document.getElementById('languageDropdown');
    
    if (languageGlobeBtn && languageDropdown) {
        // 清除可能存在的旧事件监听器
        const newLanguageGlobeBtn = languageGlobeBtn.cloneNode(true);
        languageGlobeBtn.parentNode.replaceChild(newLanguageGlobeBtn, languageGlobeBtn);
        
        // 重新获取元素引用
        const currentLanguageGlobeBtn = document.getElementById('languageGlobeBtn');
        const currentLanguageDropdown = document.getElementById('languageDropdown');
        
        // 点击地球图标显示/隐藏下拉框
        currentLanguageGlobeBtn.addEventListener('click', (event) => {
            event.stopPropagation();
            currentLanguageDropdown.classList.toggle('active');
            console.log('语言下拉框状态切换');
        });
        
        // 点击语言选项
        const languageOptions = currentLanguageDropdown.querySelectorAll('.language-option');
        languageOptions.forEach(option => {
            // 清除可能存在的旧事件监听器
            const newOption = option.cloneNode(true);
            option.parentNode.replaceChild(newOption, option);
            
            newOption.addEventListener('click', (event) => {
                const selectedLang = event.target.getAttribute('data-lang');
                console.log('语言选择改变:', selectedLang);
                handleLanguageChange(selectedLang);
                currentLanguageDropdown.classList.remove('active');
            });
        });
        
        console.log('  - 地球图标语言选择器已绑定');
    }
    
    // 设置按钮
    const settingsBtn = document.getElementById('settingsBtn');
    const apiKeySection = document.getElementById('apiKeySection');
    
    if (settingsBtn && apiKeySection) {
        // 清除可能存在的旧事件监听器
        const newSettingsBtn = settingsBtn.cloneNode(true);
        settingsBtn.parentNode.replaceChild(newSettingsBtn, settingsBtn);
        
        // 重新获取元素引用
        const currentSettingsBtn = document.getElementById('settingsBtn');
        
        currentSettingsBtn.addEventListener('click', () => {
            const isVisible = apiKeySection.style.display !== 'none';
            apiKeySection.style.display = isVisible ? 'none' : 'block';
            currentSettingsBtn.classList.toggle('active', !isVisible);
            console.log('API Key设置区域状态切换:', !isVisible);
        });
        console.log('  - 设置按钮已绑定');
    }
    
    // 点击其他地方关闭语言下拉框
    document.addEventListener('click', (event) => {
        const currentLanguageDropdown = document.getElementById('languageDropdown');
        const currentLanguageGlobeBtn = document.getElementById('languageGlobeBtn');
        
        if (currentLanguageDropdown && !currentLanguageDropdown.contains(event.target) && event.target !== currentLanguageGlobeBtn) {
            currentLanguageDropdown.classList.remove('active');
        }
    });
    
    // 翻译模态框事件
    const translateModalClose = document.getElementById('translateModalClose');
    if (translateModalClose) {
        translateModalClose.addEventListener('click', () => {
            hideTranslateModal();
        });
        console.log('  - 翻译模态框关闭按钮已绑定');
    }
    
    // 翻译语言按钮
    const translateLangBtns = document.querySelectorAll('.translate-lang-btn');
    translateLangBtns.forEach(btn => {
        btn.addEventListener('click', (event) => {
            const targetLang = event.target.getAttribute('data-target-lang');
            handleTranslateLanguageSelection(targetLang);
        });
    });
    console.log(`  - ${translateLangBtns.length} 个翻译语言按钮已绑定`);
    
    // 点击模态框背景关闭
    const translateModal = document.getElementById('translateModal');
    if (translateModal) {
        translateModal.addEventListener('click', (event) => {
            if (event.target === translateModal) {
                hideTranslateModal();
            }
        });
        console.log('  - 翻译模态框背景点击已绑定');
    }
    

    
    // API Key toggle按钮已删除，无需绑定
    
    // API Key输入框焦点事件
    const apiKeyInput = document.getElementById('apiKeyInput');
    if (apiKeyInput) {
        apiKeyInput.addEventListener('focus', () => {
            unmaskApiKey();
        });
        
        apiKeyInput.addEventListener('blur', () => {
            setTimeout(() => {
                maskApiKey();
            }, 100);
        });
        console.log('  - API Key输入框焦点事件已绑定');
    }
    
    // 返回按钮
    const backBtn = document.getElementById('backBtn');
    if (backBtn) {
        backBtn.addEventListener('click', () => {
            showMainInterface();
        });
        console.log('  - 返回按钮已绑定');
    }
    
    console.log('事件监听器绑定完成！');
}

function handleToolSelection(event) {
    console.log('handleToolSelection 被调用');
    console.log('点击的元素:', event.target);
    console.log('元素内容:', event.target.textContent);
    
    try {
        const newTool = event.target.getAttribute('data-tool');
        console.log('选择的工具:', newTool);
        console.log('之前的工具:', currentTool);
        
        // 如果是翻译工具，显示语言选择模态框
        if (newTool === 'translate') {
            showTranslateModal();
            return; // 不直接设置工具，等用户选择目标语言后再设置
        }
        
        // 检查工具是否需要用户输入
        const toolConfig = AI_TOOLS[newTool];
        if (toolConfig && toolConfig.needsInput) {
            // 直接设置工具并显示相应的输入框
            currentTool = newTool;
            
            // 更新选中状态
            document.querySelectorAll('.ai-tool-btn').forEach(btn => {
                btn.classList.remove('selected');
            });
            event.target.classList.add('selected');
            
            // 根据工具类型显示相应的输入框
            if (newTool === 'enterprise') {
                showEnterpriseInput();
            } else {
                showCustomInput();
            }
            
            updateUI();
            return;
        }
        
        // 更新选中状态
        document.querySelectorAll('.ai-tool-btn').forEach(btn => {
            btn.classList.remove('selected');
        });
        event.target.classList.add('selected');
        
        // 如果不是翻译工具，重置翻译按钮文本并清除选择的翻译语言
        if (newTool !== 'translate') {
            selectedTranslateLanguage = null;
            resetTranslateButtonText();
        }
        
        // 清理之前工具的状态
        if (newTool !== 'enterprise') {
            resetEnterpriseButtonText();
            hideEnterpriseInput();
        }
        
        // 隐藏自定义输入框（如果当前工具不需要输入）
        const newToolConfig = AI_TOOLS[newTool];
        if (!newToolConfig || !newToolConfig.needsInput || newTool === 'enterprise') {
            hideCustomInput();
        }
        
        currentTool = newTool;
        
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
            console.log('开始处理按钮被点击！');
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
            const processingText = getProcessingText();
            startBtn.innerHTML = `<span>⏳</span><span>${processingText}</span>`;
        }
        
        // 清除之前的消息
        clearMessages();
        
        // 第一步：验证是否选择了技能
        if (!currentTool) {
            throw new Error('请先选择一个技能');
        }
        
        // 第二步：验证API Key
        const apiKeyValidation = validateApiKey();
        if (!apiKeyValidation.valid) {
            // 验证失败时不跳转，直接抛出错误
            throw new Error(apiKeyValidation.message);
        }
        console.log('✅ API Key验证通过');
        
        // 第三步：立即跳转到结果界面并显示开始状态
        showResultInterface();
        showLoading('📋 正在获取Word内容...');
        
        // 第四步：获取Word内容
        console.log('📋 正在获取Word内容...');
        const content = await getWordContent();
        console.log('📋 获取到的内容:', content);
        console.log('📋 内容长度:', content.length);
        
        if (!content || content.length === 0) {
            throw new Error(`未找到内容。请先${currentContentSource === 'selection' ? '选择一些文本' : '在文档中添加内容'}。`);
        }
        
        // 在控制台显示技术信息
        console.log(`成功获取${currentContentSource === 'selection' ? '选中文本' : '文档内容'}: ${content.length} 个字符`);
        
        // 第五步：获取用户输入
        const userInput = getUserInput();
        console.log('📋 用户输入:', userInput);
        
        // 如果是翻译工具但没有选择目标语言，提示用户
        if (currentTool === 'translate' && !selectedTranslateLanguage) {
            throw new Error('请先选择翻译目标语言');
        }
        

        
        // 第六步：构建提示词
        // 特殊处理：审批建议功能根据内容源选择不同的工具
        let actualTool = currentTool;
        if (currentTool === 'academic') {
            if (currentContentSource === 'selection') {
                actualTool = 'academic'; // 选中文本使用academic (NO.003)
                console.log('📋 审批建议 - 选中文本，使用academic工具 (NO.003)');
            } else {
                actualTool = 'academicDocument'; // 整个文档使用academicDocument (NO.010)
                console.log('📋 审批建议 - 整个文档，使用academicDocument工具 (NO.010)');
            }
        }
        
        const prompt = buildPromptWithTool(content, userInput, actualTool);
        console.log('📋 构建的提示词:', prompt);
        
                        showLoading('AI正在处理中...');
        
        // 第七步：调用API
        console.log('📋 开始调用API...');
        const response = await callConversationAPI(prompt, true); // true表示新对话
        console.log('📋 API响应:', response);
        
        if (!response || response.length === 0) {
            throw new Error('AI返回了空响应');
        }
        
        showLoading('✨ 正在准备结果...');
        
        // 第八步：显示结果
        console.log('开始显示AI响应结果...');
        try {
            displayResult(response);
            console.log(`AI处理完成，生成结果: ${response.length} 个字符`);
            
            // 处理完成后跳转到结果界面
            showResultInterface();
        } catch (displayError) {
            console.error('❌ 显示结果时出错:', displayError);
            // 即使显示失败，也要保存结果
            currentResult = response;
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
        
        // 如果是API Key验证失败或其他早期错误，返回主界面
        if (error.message.includes('API Key') || error.message.includes('请先选择') || error.message.includes('请先在')) {
            showMainInterface();
        }
        
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
            const startText = LANGUAGE_TEXTS[currentLanguage]['btn.start'] || '开始处理';
                            startBtn.innerHTML = `<span>${startText}</span>`;
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
    return buildPromptWithTool(content, userInput, currentTool);
}

function buildPromptWithTool(content, userInput, toolName) {
    const tool = AI_TOOLS[toolName];
    
    if (!tool) {
        console.error('未找到工具:', toolName);
        return content; // 如果工具不存在，返回原始内容
    }
    
    let prompt = tool.prompt;
    
    // 替换模板变量
    prompt = prompt.replace('{content}', content);
    
    // 为需要用户输入的工具提供默认值
    let finalUserInput = userInput || '';
    if (tool.needsInput && !finalUserInput) {
        // 根据工具类型提供默认的用户输入
        const defaultInputs = {
            'summary': '请总结以下内容的要点',
            'grammar': '请修正以下内容的语法和表达',
            'enterprise': '请分析以下内容'
        };
        finalUserInput = defaultInputs[toolName] || '请处理以下内容';
    }
    
    prompt = prompt.replace('{userInput}', finalUserInput);
    
    // 使用当前选择的语言替换语言占位符
    const currentLanguageName = getLanguageNameForPrompt(currentLanguage);
    prompt = prompt.replace('{language}', currentLanguageName);
    
    // 如果是翻译工具，处理目标语言
    if (toolName === 'translate' && selectedTranslateLanguage) {
        const targetLanguageName = getTargetLanguageName(selectedTranslateLanguage);
        prompt = prompt.replace('{targetLanguage}', targetLanguageName);
        console.log('翻译目标语言:', targetLanguageName);
    }
    
    console.log('构建的提示词 (界面语言: ' + currentLanguage + ', 翻译目标: ' + (selectedTranslateLanguage || 'N/A') + '):', prompt);
    
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
        console.log('🚀 开始API调用:', prompt);
        
        // 使用直接API调用
        if (typeof window.directAPI !== 'undefined') {
            console.log('✨ 使用直接API调用...');
            
            if (isNewConversation) {
                // 重置对话历史
                conversationHistory = [];
                currentConversationId = null;
                
                // 使用简化的API调用
                const reply = await window.directAPI.processMessage(prompt);
                
                // 添加到对话历史
                conversationHistory.push({
                    role: 'user',
                    content: prompt
                });
                conversationHistory.push({
                    role: 'assistant',
                    content: reply
                });
                
                return reply;
            } else {
                // 继续现有对话
                if (!currentConversationId) {
                    throw new Error('缺少对话ID，请重新开始对话');
                }
                
                const reply = await window.directAPI.sendMessage(currentConversationId, prompt);
                
                // 添加到对话历史
                conversationHistory.push({
                    role: 'user',
                    content: prompt
                });
                conversationHistory.push({
                    role: 'assistant',
                    content: reply
                });
                
                return reply;
            }
        }
        
        // 备用：原始API调用逻辑
        console.log('🔄 使用备用API调用...');
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
        console.error('❌ API调用失败:', error);
        
        // 抛出错误让上层函数处理
        throw new Error(`API调用失败: ${error.message}`);
    }
}

async function handleInsert() {
            console.log('插入按钮被点击');
            console.log('当前结果长度:', currentResult ? currentResult.length : 0);
    
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
            console.log('审批建议功能：强制使用批注模式');
        }
        
        console.log('插入类型:', insertType);
        
                    showLoading('正在将内容插入Word文档...');
        
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
            insertBtn.innerHTML = '<span>插入文档</span>';
        }
        hideLoading();
    }
}

async function insertToWordWithType(text, insertType) {
            console.log('insertToWordWithType: 开始插入文本');
            console.log('要插入的文本长度:', text.length);
            console.log('插入类型:', insertType);
    
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                
                switch (insertType) {
                    case 'replace':
                        console.log('执行替换选中文本操作');
                        // 替换选中的文本
                        const selection = context.document.getSelection();
                        selection.insertText(text, Word.InsertLocation.replace);
                        break;
                        
                    case 'append':
                        console.log('执行追加到文档末尾操作');
                        // 追加到文档末尾
                        const body = context.document.body;
                        body.insertParagraph('\n' + text, Word.InsertLocation.end);
                        break;
                        
                    case 'cursor':
                        console.log('执行在光标位置插入操作');
                        // 在光标位置插入
                        const range = context.document.getSelection();
                        range.insertText(text, Word.InsertLocation.after);
                        break;
                        
                    case 'comment':
                        console.log('执行生成批注操作');
                        // 为选中文本添加批注
                        const selectionForComment = context.document.getSelection();
                        selectionForComment.load('isEmpty');
                        await context.sync();
                        
                        if (selectionForComment.isEmpty) {
                            console.log('没有选中文本，将在文档末尾插入批注内容');
                            // 如果没有选中文本，在文档末尾插入内容
                            const body = context.document.body;
                            body.insertParagraph('\n【审批建议】\n' + text, Word.InsertLocation.end);
                        } else {
                            console.log('为选中文本添加批注');
                            // 添加批注
                            selectionForComment.insertComment(text);
                        }
                        break;
                        
                    default:
                        throw new Error(`未知的插入类型: ${insertType}`);
                }
                
                console.log('正在同步到Word...');
    await context.sync();
                console.log('插入完成！');
                
                resolve();
            } catch (error) {
                console.error('插入到Word时出错:', error);
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
            console.log('开始清空操作...');
    
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
        console.log('开始显示结果，长度:', result ? result.length : 0);
        
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
        
        console.log('结果显示完成');
        
    } catch (error) {
        console.error('❌ 显示结果时出错:', error);
        console.error('错误堆栈:', error.stack);
        
        // 降级处理：直接在控制台显示结果
        console.log('降级处理 - 结果内容:', result);
    }
}

// 帮助函数：创建加载动画HTML
function createLoadingHTML(message) {
    return `
        <div class="loading-animation">
            <div class="loading-spinner"></div>
        </div>
    `;
}

function showLoading(message) {
    const resultBox = document.getElementById('resultBox');
    
    // 创建简化的加载动画
    resultBox.innerHTML = createLoadingHTML();
    resultBox.classList.add('loading');
    
    // 禁用按钮（startBtn和continueBtn不存在，跳过）
    console.log('跳过禁用不存在的按钮');
    
    console.log('🔄 显示加载状态');
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
        
        console.log('用户提示:', message);
    } catch (error) {
        console.warn('显示用户友好消息时出错:', error);
        console.log('用户提示:', message);
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
        
        // 更新企业数据输入框显示
        if (currentTool === 'enterprise') {
            showEnterpriseInput();
        } else {
            hideEnterpriseInput();
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

// 显示企业数据输入框
function showEnterpriseInput() {
    const container = document.getElementById('enterpriseInputContainer');
    if (container) {
        container.classList.remove('hidden');
        
        // 聚焦到输入框
        const textarea = document.getElementById('enterpriseInputTextarea');
        if (textarea) {
            setTimeout(() => {
                textarea.focus();
            }, 100);
        }
    }
}

// 隐藏企业数据输入框
function hideEnterpriseInput() {
    const container = document.getElementById('enterpriseInputContainer');
    if (container) {
        container.classList.add('hidden');
    }
}

// 获取用户输入
function getUserInput() {
    // 检查当前工具是否需要用户输入
    const toolConfig = AI_TOOLS[currentTool];
    if (!toolConfig || !toolConfig.needsInput) {
        return '';
    }
    
    // 根据工具类型获取相应的输入
    if (currentTool === 'enterprise') {
        const textarea = document.getElementById('enterpriseInputTextarea');
        if (textarea) {
            return textarea.value.trim();
        }
    } else {
        // 总结摘要、语法修正等使用自定义输入框
        const textarea = document.getElementById('customInputTextarea');
        if (textarea) {
            return textarea.value.trim();
        }
    }
    
    return '';
}

// 语言处理相关函数
function handleLanguageChange(language) {
    console.log('切换语言:', language);
    currentLanguage = language;
    updateLanguageDisplay();
}

function updateLanguageDisplay() {
    const texts = LANGUAGE_TEXTS[currentLanguage];
    if (!texts) {
        console.error('未找到语言配置:', currentLanguage);
        return;
    }
    
    console.log('更新界面语言为:', currentLanguage);
    
    // 更新所有带有data-i18n属性的元素
    document.querySelectorAll('[data-i18n]').forEach(element => {
        const key = element.getAttribute('data-i18n');
        if (texts[key]) {
            element.textContent = texts[key];
        }
    });
    
    // 更新所有带有 data-i18n-placeholder 属性的元素
    document.querySelectorAll('[data-i18n-placeholder]').forEach(element => {
        const key = element.getAttribute('data-i18n-placeholder');
        if (texts[key]) {
            element.placeholder = texts[key];
        }
    });
    
    // 更新placeholder
    document.querySelectorAll('[data-i18n-placeholder]').forEach(element => {
        const key = element.getAttribute('data-i18n-placeholder');
        if (texts[key]) {
            element.placeholder = texts[key];
        }
    });
    
    // 如果翻译工具已选择目标语言，更新翻译按钮文本
    // 更新翻译按钮文本
    const translateBtn = document.querySelector('[data-tool="translate"]');
    if (translateBtn) {
        if (currentTool === 'translate' && selectedTranslateLanguage) {
            updateTranslateButtonText(translateBtn, selectedTranslateLanguage);
        } else {
            // 如果没有选择翻译目标语言，显示默认文本
            resetTranslateButtonText();
        }
    }
    

    
    console.log('界面语言更新完成');
}

function getLanguageNameForPrompt(languageCode) {
    const languageMap = {
        'zh-cn': '中文',
        'en': 'English',
        'th': 'Thai',
        'ja': '日本語',
        'zh-tw': '繁體中文'
    };
    return languageMap[languageCode] || '中文';
}

function getProcessingText() {
    const processingTexts = {
        'zh-cn': '处理中...',
        'en': 'Processing...',
        'th': 'กำลังประมวลผล...',
        'ja': '処理中...',
        'zh-tw': '處理中...'
    };
    return processingTexts[currentLanguage] || '处理中...';
}

// 翻译功能相关函数
function showTranslateModal() {
    const modal = document.getElementById('translateModal');
    if (modal) {
        modal.style.display = 'flex';
        console.log('显示翻译语言选择模态框');
    }
}

function hideTranslateModal() {
    const modal = document.getElementById('translateModal');
    if (modal) {
        modal.style.display = 'none';
        console.log('隐藏翻译语言选择模态框');
    }
}

function handleTranslateLanguageSelection(targetLang) {
    console.log('选择翻译目标语言:', targetLang);
    selectedTranslateLanguage = targetLang;
    
    // 设置翻译工具为当前工具
    currentTool = 'translate';
    
    // 更新按钮选中状态和显示文本
    document.querySelectorAll('.ai-tool-btn').forEach(btn => {
        btn.classList.remove('selected');
    });
    const translateBtn = document.querySelector('[data-tool="translate"]');
    if (translateBtn) {
        translateBtn.classList.add('selected');
        // 更新按钮显示文本
        updateTranslateButtonText(translateBtn, targetLang);
    }
    
    // 隐藏自定义输入框
    hideCustomInput();
    
    // 更新UI状态
    updateUI();
    
    // 隐藏模态框
    hideTranslateModal();
    
    console.log(`翻译工具设置完成，目标语言: ${targetLang}`);
}

function updateTranslateButtonText(button, targetLang) {
    const texts = LANGUAGE_TEXTS[currentLanguage];
    if (!texts || !targetLang) return;
    
    const translatePrefix = texts['translate.to'] || '翻译成';
    const targetLanguageName = getTargetLanguageName(targetLang);
    
    button.textContent = `${translatePrefix} ${targetLanguageName}`;
    
    console.log(`翻译按钮文本已更新: ${button.textContent}`);
}

function resetTranslateButtonText() {
    const translateBtn = document.querySelector('[data-tool="translate"]');
    if (translateBtn) {
        const texts = LANGUAGE_TEXTS[currentLanguage];
        const baseText = texts ? texts['btn.translate'] : '翻译';
        translateBtn.textContent = baseText;
        console.log(`翻译按钮文本已重置: ${translateBtn.textContent}`);
    }
}

function getTargetLanguageName(langCode) {
    const languageNames = {
        'en': 'English',
        'zh-cn': '中文',
        'zh-tw': '繁體中文',
        'ja': '日本語',
        'ko': '한국어',
        'th': 'ไทย',
        'fr': 'Français',
        'de': 'Deutsch',
        'es': 'Español'
    };
    return languageNames[langCode] || langCode;
}

// 企业数据模态框相关功能
function resetEnterpriseButtonText() {
    const enterpriseBtn = document.querySelector('[data-tool="enterprise"]');
    if (enterpriseBtn) {
        const texts = LANGUAGE_TEXTS[currentLanguage];
        enterpriseBtn.textContent = texts['btn.enterprise'] || '企业数据';
        console.log('企业数据按钮文本已重置');
    }
}

// 界面切换功能
function showMainInterface() {
    const mainInterface = document.getElementById('mainInterface');
    const resultInterface = document.getElementById('resultInterface');
    
    if (mainInterface && resultInterface) {
        mainInterface.style.display = 'block';
        resultInterface.style.display = 'none';
        console.log('切换到主界面');
    }
}

function showResultInterface() {
    const mainInterface = document.getElementById('mainInterface');
    const resultInterface = document.getElementById('resultInterface');
    
    if (mainInterface && resultInterface) {
        mainInterface.style.display = 'none';
        resultInterface.style.display = 'block';
        console.log('切换到结果界面');
    }
}

// API Key 相关功能
function maskApiKey() {
    const apiKeyInput = document.getElementById('apiKeyInput');
    if (apiKeyInput && apiKeyInput.value) {
        const value = apiKeyInput.value;
        if (value.length > 8) {
            // 显示前4个字符和后4个字符，中间用星号替换
            const start = value.substring(0, 4);
            const end = value.substring(value.length - 4);
            const middle = '*'.repeat(value.length - 8);
            apiKeyInput.setAttribute('data-original-value', value);
            apiKeyInput.value = start + middle + end;
        }
    }
}

function unmaskApiKey() {
    const apiKeyInput = document.getElementById('apiKeyInput');
    if (apiKeyInput) {
        const originalValue = apiKeyInput.getAttribute('data-original-value');
        if (originalValue) {
            apiKeyInput.value = originalValue;
        }
    }
}

function getApiKey() {
    const apiKeyInput = document.getElementById('apiKeyInput');
    if (apiKeyInput) {
        // 如果有原始值，返回原始值，否则返回当前值
        const originalValue = apiKeyInput.getAttribute('data-original-value');
        return originalValue ? originalValue.trim() : apiKeyInput.value.trim();
    }
    return '';
}

function validateApiKey() {
    const apiKey = getApiKey();
    
    if (!apiKey) {
        return { valid: false, message: '请输入API Key' };
    }
    
    // 基本格式验证 - 应该以 app- 开头，且有足够的长度
    if (!apiKey.startsWith('app-') || apiKey.length < 20) {
        return { valid: false, message: 'API Key格式不正确，应该以app-开头且有足够长度' };
    }
    
    return { valid: true, message: 'API Key格式正确' };
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
                console.log('文本长度:', selection.text.length);
                
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
