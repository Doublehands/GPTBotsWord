from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import os
import signal
import atexit

app = Flask(__name__)
CORS(app, resources={
    r"/*": {
        "origins": [
            "http://localhost:3000",
            "https://localhost:3000",
            "http://localhost:3001",
            "https://localhost:3001",
            "https://localhost",
            "https://127.0.0.1",
            "null",  # 允许来自file://协议的请求
        ],
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": ["Content-Type", "Authorization", "Accept", "Origin", "X-Requested-With"]
    }
})

API_KEY = "app-cqAuvC3vC7d7LybynuoZdK9D"
API_BASE_URL = "https://api.gptbots.ai"

# 创建PID文件
def create_pid_file():
    pid = os.getpid()
    with open("flask-server.pid", "w") as f:
        f.write(str(pid))
    print(f"已创建PID文件: {os.path.abspath('flask-server.pid')}")

# 清理PID文件
def cleanup_pid_file():
    try:
        os.remove("flask-server.pid")
        print("已清理PID文件")
    except:
        pass

# 注册清理函数
atexit.register(cleanup_pid_file)

# 健康检查端点
@app.route('/')
def health_check():
    return jsonify({"status": "ok", "message": "本地代理服务器运行正常"})

# 创建对话
@app.route('/api/v1/conversation', methods=['POST', 'OPTIONS'])
def create_conversation():
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        # 获取请求数据
        data = request.get_json()
        print(f"\n📝 创建对话请求:")
        print(f"请求体: {data}")
        
        # 转发到GPTBots API
        headers = {
            'Authorization': f'Bearer {API_KEY}',
            'Content-Type': 'application/json',
            'User-Agent': 'GPTBots-Word-Addin/1.0',
            'Accept': 'application/json'
        }
        
        response = requests.post(
            f"{API_BASE_URL}/api/v1/conversation",
            json=data,
            headers=headers
        )
        
        print(f"API响应状态码: {response.status_code}")
        print(f"API响应内容: {response.text}\n")
        
        return response.text, response.status_code, {'Content-Type': 'application/json'}
        
    except Exception as e:
        print(f"❌ 创建对话失败: {str(e)}")
        return jsonify({
            "error": "创建对话失败",
            "message": str(e)
        }), 500

# 发送消息
@app.route('/api/v2/conversation/message', methods=['POST', 'OPTIONS'])
def send_message():
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        # 获取请求数据
        data = request.get_json()
        print(f"\n💬 发送消息请求:")
        print(f"请求体: {data}")
        
        # 转发到GPTBots API
        headers = {
            'Authorization': f'Bearer {API_KEY}',
            'Content-Type': 'application/json',
            'User-Agent': 'GPTBots-Word-Addin/1.0',
            'Accept': 'application/json'
        }
        
        response = requests.post(
            f"{API_BASE_URL}/api/v2/conversation/message",
            json=data,
            headers=headers
        )
        
        print(f"API响应状态码: {response.status_code}")
        print(f"API响应内容: {response.text}\n")
        
        return response.text, response.status_code, {'Content-Type': 'application/json'}
        
    except Exception as e:
        print(f"❌ 发送消息失败: {str(e)}")
        return jsonify({
            "error": "发送消息失败",
            "message": str(e)
        }), 500

if __name__ == '__main__':
    try:
        # 创建PID文件
        create_pid_file()
        
        print("\n🚀 启动本地代理服务器...")
        print("✅ 本地代理服务器已启动")
        print("🌐 服务器地址: http://localhost:3001")
        print("\n可用端点:")
        print("📝 创建对话: http://localhost:3001/api/v1/conversation")
        print("💬 发送消息: http://localhost:3001/api/v2/conversation/message")
        print("\n按 Ctrl+C 停止服务器")
        
        # 启动Flask服务器
        app.run(host='localhost', port=3001, debug=True)
        
    except KeyboardInterrupt:
        print("\n正在停止服务器...")
        cleanup_pid_file() 