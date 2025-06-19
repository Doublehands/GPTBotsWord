# Word GPT Plus API 配置指南

本指南将帮助您配置 Word GPT Plus 插件，使其能够与您的对话Agent API正常工作。

## 快速开始

### 1. 基本配置

打开 `src/taskpane/api-config.js` 文件，修改以下配置：

```javascript
const API_CONFIG = {
    // 您的API基础URL
    baseUrl: 'https://your-api-endpoint.com/api',
    
    // 聊天API端点 (相对于baseUrl)
    chatEndpoint: '/chat',
    
    // 请求头配置
    headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer YOUR_API_KEY', // 取消注释并设置您的API密钥
    },
};
```

### 2. 使用预设配置

如果您的API遵循常见格式，可以使用预设配置：

#### OpenAI 格式 API
```javascript
// 在 taskpane.js 中调用
applyPreset('openai');
// 然后设置您的API密钥
API_CONFIG.headers['Authorization'] = 'Bearer YOUR_OPENAI_API_KEY';
```

#### Azure OpenAI 格式 API
```javascript
// 在 taskpane.js 中调用
applyPreset('azure');
// 设置您的Azure配置
API_CONFIG.baseUrl = 'https://your-resource.openai.azure.com';
API_CONFIG.headers['api-key'] = 'YOUR_AZURE_API_KEY';
// 修改部署名称
API_CONFIG.chatEndpoint = '/openai/deployments/YOUR_DEPLOYMENT_NAME/chat/completions?api-version=2023-05-15';
```

## 详细配置说明

### API响应格式配置

如果您的API返回格式与常见格式不同，请修改 `responseMapping` 对象：

```javascript
responseMapping: {
    message: 'data.result.text',    // 您的API中消息内容的字段路径
    error: 'error.message',         // 您的API中错误信息的字段路径
    status: 'status.code'           // 您的API中状态码的字段路径
}
```

支持的路径格式：
- 简单字段：`'message'`
- 嵌套字段：`'data.result.text'`
- 数组索引：`'choices[0].message.content'`

### 常见API格式示例

#### 1. 标准ChatGPT格式
```json
{
  "choices": [
    {
      "message": {
        "content": "这是AI的回复内容"
      }
    }
  ]
}
```
配置：
```javascript
responseMapping: {
    message: 'choices[0].message.content'
}
```

#### 2. 简单格式
```json
{
  "message": "这是AI的回复内容",
  "status": "success"
}
```
配置：
```javascript
responseMapping: {
    message: 'message',
    status: 'status'
}
```

#### 3. 嵌套格式
```json
{
  "data": {
    "result": {
      "text": "这是AI的回复内容"
    }
  },
  "code": 200
}
```
配置：
```javascript
responseMapping: {
    message: 'data.result.text',
    status: 'code'
}
```

### 请求参数配置

您可以修改默认的请求参数：

```javascript
defaultParams: {
    model: 'your-model-name',     // 模型名称
    temperature: 0.7,             // 温度参数
    max_tokens: 2000,             // 最大tokens
    top_p: 1,                     // top_p参数
    frequency_penalty: 0,         // 频率惩罚
    presence_penalty: 0,          // 存在惩罚
}
```

## 测试配置

### 1. 在浏览器控制台测试

在Word中打开插件后，按F12打开开发者工具，在控制台中输入：

```javascript
// 检查配置
console.log('API配置:', API_CONFIG);
console.log('完整API URL:', getFullApiUrl());

// 测试API连接 (可选)
fetch(getFullApiUrl(), {
    method: 'POST',
    headers: API_CONFIG.headers,
    body: JSON.stringify(buildRequestData([
        {role: 'user', content: '你好'}
    ]))
}).then(response => response.json()).then(data => {
    console.log('API响应:', data);
    console.log('解析结果:', parseApiResponse(data));
});
```

### 2. 使用插件测试

1. 在Word文档中输入一些文本
2. 选中文本
3. 在插件中输入问题
4. 点击"开始处理"按钮
5. 查看控制台是否有错误信息

## 常见问题解决

### 1. CORS 错误
如果遇到跨域问题，需要在您的API服务器上配置CORS头：
```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: POST, GET, OPTIONS
Access-Control-Allow-Headers: Content-Type, Authorization
```

### 2. 认证失败
检查API密钥是否正确设置：
```javascript
// 检查认证头是否正确
console.log('认证头:', API_CONFIG.headers);
```

### 3. 响应解析失败
检查响应格式映射是否正确：
```javascript
// 打印实际API响应
console.log('实际响应:', response);
// 检查解析结果
console.log('解析结果:', parseApiResponse(response));
```

### 4. 超时问题
增加超时时间：
```javascript
API_CONFIG.timeout = 60000; // 60秒
```

## 高级配置

### 动态配置
您可以在运行时动态修改配置：

```javascript
// 在 taskpane.js 的 initializeApp() 函数中
function initializeApp() {
    // 根据用户设置或环境变量动态配置
    if (localStorage.getItem('apiUrl')) {
        API_CONFIG.baseUrl = localStorage.getItem('apiUrl');
    }
    
    if (localStorage.getItem('apiKey')) {
        API_CONFIG.headers['Authorization'] = `Bearer ${localStorage.getItem('apiKey')}`;
    }
    
    // 其他初始化代码...
}
```

### 多种API支持
您可以支持多种API格式：

```javascript
// 根据用户选择切换API格式
function switchApiFormat(format) {
    switch(format) {
        case 'openai':
            applyPreset('openai');
            break;
        case 'azure':
            applyPreset('azure');
            break;
        default:
            applyPreset('generic');
    }
}
```

## 支持

如果您在配置过程中遇到问题，请：

1. 检查浏览器控制台的错误信息
2. 确认API端点是否可访问
3. 验证API密钥是否有效
4. 检查请求和响应格式是否匹配

您也可以通过修改 `console.log` 语句来增加调试信息，以便更好地理解数据流。 