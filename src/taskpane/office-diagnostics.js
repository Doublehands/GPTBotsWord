// Office.js 诊断工具

window.officeDiagnostics = {
    // 完整的Office.js状态检查
    checkOfficeStatus: function() {
        console.log('🔍 === Office.js 完整诊断 ===');
        
        // 1. 检查基础对象
        console.log('1️⃣ 基础对象检查:');
        console.log('- Office对象:', typeof Office !== 'undefined' ? '✅ 已加载' : '❌ 未加载');
        console.log('- Word对象:', typeof Word !== 'undefined' ? '✅ 已加载' : '❌ 未加载');
        console.log('- OfficeExtension:', typeof OfficeExtension !== 'undefined' ? '✅ 已加载' : '❌ 未加载');
        
        // 2. 检查Office上下文
        if (typeof Office !== 'undefined') {
            console.log('2️⃣ Office上下文信息:');
            try {
                console.log('- 主机应用:', Office.context?.host || '未知');
                console.log('- 平台:', Office.context?.platform || '未知');
                console.log('- 版本:', Office.context?.requirements?.sets || '未知');
                console.log('- 语言:', Office.context?.displayLanguage || '未知');
            } catch (e) {
                console.log('- 上下文信息获取失败:', e.message);
            }
        }
        
        // 3. 检查URL和环境
        console.log('3️⃣ 环境检查:');
        console.log('- 当前URL:', window.location.href);
        console.log('- 是否HTTPS:', window.location.protocol === 'https:');
        console.log('- 用户代理:', navigator.userAgent.includes('Microsoft Office') ? '✅ Office环境' : '❌ 非Office环境');
        
        // 4. 检查脚本加载
        console.log('4️⃣ 脚本加载检查:');
        const scripts = Array.from(document.querySelectorAll('script'));
        const officeScript = scripts.find(s => s.src.includes('office.js'));
        console.log('- Office.js脚本:', officeScript ? '✅ 已加载' : '❌ 未找到');
        if (officeScript) {
            console.log('- 脚本源:', officeScript.src);
        }
        
        return {
            officeLoaded: typeof Office !== 'undefined',
            wordLoaded: typeof Word !== 'undefined',
            context: typeof Office !== 'undefined' ? Office.context : null
        };
    },
    
    // 测试Word连接
    testWordConnection: async function() {
        console.log('🔗 测试Word连接...');
        
        if (typeof Office === 'undefined') {
            console.error('❌ Office对象未定义');
            return { success: false, error: 'Office对象未定义' };
        }
        
        if (typeof Word === 'undefined') {
            console.error('❌ Word对象未定义');
            return { success: false, error: 'Word对象未定义' };
        }
        
        try {
            const result = await new Promise((resolve, reject) => {
                Word.run(async (context) => {
                    try {
                        // 基础连接测试
                        const doc = context.document;
                        doc.load('title');
                        
                        const selection = context.document.getSelection();
                        selection.load('text');
                        
                        const body = context.document.body;
                        body.load('text');
                        
                        await context.sync();
                        
                        resolve({
                            success: true,
                            title: doc.title,
                            selectionText: selection.text,
                            documentLength: body.text.length
                        });
                    } catch (error) {
                        reject(error);
                    }
                });
            });
            
            console.log('✅ Word连接成功!');
            console.log('- 文档标题:', result.title || '无标题');
            console.log('- 选中文本:', result.selectionText || '无选中');
            console.log('- 文档长度:', result.documentLength);
            
            return result;
            
        } catch (error) {
            console.error('❌ Word连接失败:', error);
            return { success: false, error: error.message };
        }
    },
    
    // 修复建议
    getSuggestions: function() {
        const status = this.checkOfficeStatus();
        const suggestions = [];
        
        if (!status.officeLoaded) {
            suggestions.push('1. 刷新页面重新加载Office.js');
            suggestions.push('2. 检查网络连接');
            suggestions.push('3. 确保在Word中运行插件');
        }
        
        if (!status.wordLoaded) {
            suggestions.push('1. 确保插件在Word中运行（不是其他Office应用）');
            suggestions.push('2. 检查Word版本是否支持Office插件');
        }
        
        if (suggestions.length === 0) {
            suggestions.push('✅ Office.js状态正常！');
        }
        
        console.log('💡 修复建议:');
        suggestions.forEach(suggestion => console.log(suggestion));
        
        return suggestions;
    }
};

// 快捷命令
window.checkOffice = window.officeDiagnostics.checkOfficeStatus;
window.testWord = window.officeDiagnostics.testWordConnection;
window.fixSuggestions = window.officeDiagnostics.getSuggestions;

console.log('🔧 Office.js诊断工具已加载!');
console.log('可用命令:');
console.log('- checkOffice(): 检查Office.js状态');
console.log('- testWord(): 测试Word连接');
console.log('- fixSuggestions(): 获取修复建议'); 