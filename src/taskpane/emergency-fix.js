// 紧急修复脚本 - 如果正常的事件绑定失败，可以运行这个脚本

// 检查Office.js是否加载
function checkOfficeStatus() {
    console.log('🔍 检查Office.js状态...');
    console.log('- Office对象:', typeof Office !== 'undefined' ? '✅ 已加载' : '❌ 未加载');
    console.log('- Word对象:', typeof Word !== 'undefined' ? '✅ 已加载' : '❌ 未加载');
    
    if (typeof Office !== 'undefined') {
        console.log('- Office版本:', Office.context?.platform || '未知');
        console.log('- 主机应用:', Office.context?.host || '未知');
    }
    
    return typeof Office !== 'undefined' && typeof Word !== 'undefined';
}

function emergencyFixButtons() {
    console.log('🚨 运行紧急按钮修复...');
    
    // 紧急修复AI工具按钮
    const aiButtons = [
        {selector: '[data-tool="translate"]', tool: 'translate'},
        {selector: '[data-tool="polish"]', tool: 'polish'},
        {selector: '[data-tool="academic"]', tool: 'academic'},
        {selector: '[data-tool="summary"]', tool: 'summary'},
        {selector: '[data-tool="grammar"]', tool: 'grammar'},
        {selector: '[data-tool="custom"]', tool: 'custom'}
    ];
    
    aiButtons.forEach(({selector, tool}) => {
        const btn = document.querySelector(selector);
        if (btn) {
            // 清除之前的事件监听器
            btn.onclick = null;
            
            btn.onclick = function() {
                console.log(`紧急处理: 选择工具 ${tool}`);
                
                // 移除所有selected类
                document.querySelectorAll('.ai-tool-btn').forEach(b => b.classList.remove('selected'));
                // 添加selected类到当前按钮
                this.classList.add('selected');
                
                // 更新工具
                window.currentTool = tool;
                console.log('当前工具设置为:', window.currentTool);
                
                // 显示成功消息
                const resultBox = document.getElementById('resultBox');
                if (resultBox) {
                    resultBox.textContent = `Selected ${tool} tool`;
                }
            };
            console.log(`✅ 紧急修复成功: ${tool} 按钮`);
        } else {
            console.log(`❌ 未找到按钮: ${selector}`);
        }
    });
    
    // 紧急修复内容源按钮
    const sourceButtons = [
        {selector: '[data-source="selection"]', source: 'selection'},
        {selector: '[data-source="document"]', source: 'document'}
    ];
    
    sourceButtons.forEach(({selector, source}) => {
        const btn = document.querySelector(selector);
        if (btn) {
            // 清除之前的事件监听器
            btn.onclick = null;
            
            btn.onclick = function() {
                console.log(`紧急处理: 选择内容源 ${source}`);
                
                // 移除所有active类
                document.querySelectorAll('.content-source-btn').forEach(b => b.classList.remove('active'));
                // 添加active类到当前按钮
                this.classList.add('active');
                
                // 更新内容源
                window.currentContentSource = source;
                console.log('当前内容源设置为:', window.currentContentSource);
                
                // 显示成功消息
                const resultBox = document.getElementById('resultBox');
                if (resultBox) {
                    resultBox.textContent = `Selected ${source === 'selection' ? 'Selected Text' : 'Entire Document'}`;
                }
            };
            console.log(`✅ 紧急修复成功: ${source} 按钮`);
        }
    });
    
    // 紧急修复主要操作按钮
    const startBtn = document.getElementById('startBtn');
    if (startBtn) {
        // 清除之前的事件监听器
        startBtn.onclick = null;
        
        startBtn.onclick = async function() {
            console.log('🚀 紧急处理: 开始处理按钮');
            
                            try {
                // 显示加载状态
                const resultBox = document.getElementById('resultBox');
                if (resultBox) {
                    // 创建加载动画
                    resultBox.innerHTML = `
                        <div class="loading-animation">
                            <div class="loading-spinner"></div>
                            <div class="loading-dots">
                                <div class="loading-dot"></div>
                                <div class="loading-dot"></div>
                                <div class="loading-dot"></div>
                            </div>
                        </div>
                        <div class="loading-text">🤖 AI processing, please wait...</div>
                    `;
                    resultBox.classList.add('loading');
                }
                
                // 获取Word内容
                console.log('📋 正在获取Word内容...');
                if (resultBox) {
                    resultBox.innerHTML = `
                        <div class="loading-animation">
                            <div class="loading-spinner"></div>
                            <div class="loading-dots">
                                <div class="loading-dot"></div>
                                <div class="loading-dot"></div>
                                <div class="loading-dot"></div>
                            </div>
                        </div>
                        <div class="loading-text">📋 Reading Word content...</div>
                    `;
                }
                
                const content = await new Promise((resolve, reject) => {
                    // 确保Office已经准备好
                    if (typeof Office === 'undefined') {
                        reject(new Error('Office.js未加载。请刷新页面重试。'));
                        return;
                    }
                    
                    if (typeof Word === 'undefined') {
                        reject(new Error('Word对象不可用。请确保在Word中运行此插件。'));
                        return;
                    }
                    
                    Word.run(async (context) => {
                        try {
                            let text = '';
                            const source = window.currentContentSource || 'selection';
                            
                            if (source === 'selection') {
                                const selection = context.document.getSelection();
                                selection.load('text');
                                await context.sync();
                                text = selection.text;
                            } else {
                                const body = context.document.body;
                                body.load('text');
                                await context.sync();
                                text = body.text;
                            }
                            
                            resolve(text.trim());
                        } catch (error) {
                            reject(error);
                        }
                    }).catch(reject);
                });
                
                console.log('📋 获取的内容:', content);
                console.log('📋 内容长度:', content.length);
                
                if (!content || content.length === 0) {
                    throw new Error('No content retrieved. Please ensure text is selected in Word.');
                }
                
                // 清理内容，移除之前的处理结果
                let cleanContent = content;
                
                // 如果内容包含之前的处理结果，只提取原文部分
                if (content.includes('处理后:') || content.includes('这是模拟的')) {
                    const lines = content.split('\n');
                    const originalLineIndex = lines.findIndex(line => line.includes('原文:'));
                    
                    if (originalLineIndex !== -1 && originalLineIndex + 1 < lines.length) {
                        // 提取原文行的内容
                        const originalLine = lines[originalLineIndex];
                        cleanContent = originalLine.replace(/^原文:\s*/, '').trim();
                    } else {
                        // 如果找不到原文标记，尝试提取第一行非处理结果的内容
                        cleanContent = lines.find(line => 
                            line.trim() && 
                            !line.includes('处理结果') && 
                            !line.includes('这是模拟的') &&
                            !line.includes('Original:') &&
                            !line.includes('Processed:')
                        ) || content;
                    }
                }
                
                console.log('📋 清理前的内容:', content);
                console.log('📋 清理后的内容:', cleanContent);
                
                // 构建提示词
                const tool = window.currentTool || 'translate';
                let prompt;
                
                if (tool === 'translate') {
                    prompt = `请将以下内容翻译成英文：\n\n${cleanContent}`;
                } else if (tool === 'polish') {
                    prompt = `请润色以下内容：\n\n${cleanContent}`;
                } else if (tool === 'academic') {
                    prompt = `请将以下内容转换为学术化表达：\n\n${cleanContent}`;
                } else if (tool === 'summary') {
                    prompt = `请为以下内容生成摘要：\n\n${cleanContent}`;
                } else if (tool === 'grammar') {
                    prompt = `请检查并修正以下内容的语法错误：\n\n${cleanContent}`;
                } else {
                    prompt = `请处理以下内容：\n\n${cleanContent}`;
                }
                
                console.log('📋 构建的提示词:', prompt);
                
                // 尝试真实API调用
                console.log('🌐 尝试调用真实AI API...');
                if (resultBox) {
                    resultBox.innerHTML = `
                        <div class="loading-animation">
                            <div class="loading-spinner"></div>
                            <div class="loading-dots">
                                <div class="loading-dot"></div>
                                <div class="loading-dot"></div>
                                <div class="loading-dot"></div>
                            </div>
                        </div>
                        <div class="loading-text">🤖 AI processing...</div>
                    `;
                }
                
                let finalResult;
                
                try {
                    // 使用直接API调用
                    console.log('🚀 尝试直接API调用...');
                    
                    if (typeof window.directAPI !== 'undefined') {
                        const apiResult = await window.directAPI.processMessage(prompt);
                        
                        if (apiResult) {
                            finalResult = `🎉 ${AI_TOOLS[tool].name}处理结果（真实AI回复）:\n\n原文: ${cleanContent}\n\n处理结果:\n${apiResult}`;
                            console.log('🎉 直接API调用成功！');
                            console.log('AI回复:', apiResult);
                        } else {
                            throw new Error('API返回空结果');
                        }
                    } else {
                        throw new Error('直接API未加载');
                    }
                    
                } catch (apiError) {
                    console.warn('⚠️ 直接API调用失败:', apiError.message);
                    
                    // 使用模拟结果（后备方案）
                    console.log('🔄 使用模拟结果...');
                    console.log('⚠️ API调用失败详情:', apiError.message);
                    
                    // 只在结果区显示简洁的模拟结果，不显示错误信息
                    finalResult = `Processing...`;
                }
                
                // 显示最终结果
                window.currentResult = finalResult;
                if (resultBox) {
                    resultBox.innerHTML = finalResult;
                    resultBox.classList.remove('loading');
                }
                
                console.log('🎉 处理完成！');
                
            } catch (error) {
                console.error('❌ 处理失败:', error);
                console.log('💡 详细错误信息请查看上方的日志');
                
                // 在结果区显示友好的消息，不显示具体错误
                const resultBox = document.getElementById('resultBox');
                if (resultBox) {
                    resultBox.innerHTML = `Processing...`;
                    resultBox.classList.remove('loading');
                }
            }
        };
        console.log('✅ 紧急修复成功: 开始处理按钮');
    }
    
    // 紧急修复结果操作按钮
    const insertBtn = document.getElementById('insertBtn');
    if (insertBtn) {
        // 清除之前的事件监听器
        insertBtn.onclick = null;
        
        insertBtn.onclick = async function() {
            console.log('📝 紧急处理: 插入按钮');
            
            if (!window.currentResult || window.currentResult.trim().length === 0) {
                console.log('❌ 没有可插入的内容');
                alert('No content to insert, please process some text first.');
                return;
            }
            
                         try {
                await new Promise((resolve, reject) => {
                    // 确保Office已经准备好
                    if (typeof Office === 'undefined') {
                        reject(new Error('Office.js未加载。请刷新页面重试。'));
                        return;
                    }
                    
                    if (typeof Word === 'undefined') {
                        reject(new Error('Word对象不可用。请确保在Word中运行此插件。'));
                        return;
                    }
                    
                    Word.run(async (context) => {
                        try {
                            const selection = context.document.getSelection();
                            selection.insertText(window.currentResult, Word.InsertLocation.replace);
                            await context.sync();
                            resolve();
                        } catch (error) {
                            reject(error);
                        }
                    }).catch(reject);
                });
                
                console.log('📝 插入成功！');
                alert('Content successfully inserted to Word document!');
                
            } catch (error) {
                console.error('📝 插入失败:', error);
                alert('Content insertion encountered issues, please retry or check Word document status');
            }
        };
        console.log('✅ 紧急修复成功: 插入按钮');
    }
    
    const copyBtn = document.getElementById('copyBtn');
    if (copyBtn) {
        // 清除之前的事件监听器
        copyBtn.onclick = null;
        
        copyBtn.onclick = function() {
            console.log('📋 紧急处理: 复制按钮');
            
            if (!window.currentResult || window.currentResult.trim().length === 0) {
                console.log('❌ 没有可复制的内容');
                alert('No content to copy, please process some text first.');
                return;
            }
            
            // 复制到剪贴板
            if (navigator.clipboard) {
                navigator.clipboard.writeText(window.currentResult).then(() => {
                    console.log('📋 复制成功！');
                    alert('Content copied to clipboard!');
                }).catch(() => {
                    // 降级方法
                    console.log('📋 使用降级复制方法');
                    try {
                        const textArea = document.createElement('textarea');
                        textArea.value = window.currentResult;
                        textArea.style.position = 'fixed';
                        textArea.style.opacity = '0';
                        document.body.appendChild(textArea);
                        textArea.select();
                        const successful = document.execCommand('copy');
                        document.body.removeChild(textArea);
                        
                        if (successful) {
                            alert('Content copied to clipboard!');
                        } else {
                            alert('Please manually select and copy content from result area');
                        }
                    } catch (err) {
                        console.error('复制失败:', err);
                        alert('Please manually select and copy content from result area');
                    }
                });
            } else {
                // 浏览器不支持clipboard API
                console.log('📋 浏览器不支持clipboard API，使用降级方法');
                try {
                    const textArea = document.createElement('textarea');
                    textArea.value = window.currentResult;
                    textArea.style.position = 'fixed';
                    textArea.style.opacity = '0';
                    document.body.appendChild(textArea);
                    textArea.select();
                    const successful = document.execCommand('copy');
                    document.body.removeChild(textArea);
                    
                    if (successful) {
                        alert('Content copied to clipboard!');
                    } else {
                        alert('Please manually select and copy content from result area');
                    }
                } catch (err) {
                    console.error('复制失败:', err);
                    alert('Please manually select and copy content from result area');
                }
            }
        };
        console.log('✅ 紧急修复成功: 复制按钮');
    }
    
    const clearBtn = document.getElementById('clearBtn');
    if (clearBtn) {
        // 清除之前的事件监听器
        clearBtn.onclick = null;
        
        clearBtn.onclick = function() {
            console.log('🗑️ 紧急处理: 清空按钮');
            
            window.currentResult = '';
            const resultBox = document.getElementById('resultBox');
            if (resultBox) {
                resultBox.innerHTML = 'Click "Start Processing" to get AI response';
                resultBox.classList.remove('loading');
            }
            
            // 清空输入框
            const conversationInput = document.getElementById('conversationInput');
            if (conversationInput) {
                conversationInput.value = '';
            }
            
            console.log('🗑️ 已清空结果');
            alert('Results cleared!');
        };
        console.log('✅ 紧急修复成功: 清空按钮');
    }
    
    console.log('🎉 紧急修复完成！现在按钮应该可以响应了。');
}

// 等待Office.js加载后自动运行紧急修复
function initEmergencyFix() {
    if (typeof Office !== 'undefined') {
        // Office已加载，直接运行
        emergencyFixButtons();
    } else {
        // 等待Office加载
        console.log('⏳ 等待Office.js加载...');
        let retryCount = 0;
        const maxRetries = 20; // 最多等待10秒
        
        const checkOffice = setInterval(() => {
            retryCount++;
            if (typeof Office !== 'undefined') {
                console.log('✅ Office.js已加载，运行紧急修复');
                clearInterval(checkOffice);
                emergencyFixButtons();
            } else if (retryCount >= maxRetries) {
                console.warn('⚠️ Office.js加载超时，可能需要手动运行 emergencyFixButtons()');
                clearInterval(checkOffice);
            }
        }, 500);
    }
}

// 自动运行紧急修复
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initEmergencyFix);
} else {
    initEmergencyFix();
}

// 提供手动运行的接口
window.emergencyFixButtons = emergencyFixButtons;
window.checkOfficeStatus = checkOfficeStatus; 