import os
import time
import base64
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class QwenVLProcessor:
    def __init__(self, api_key, watch_folder):
        self.API_KEY = api_key  # API key is passed as a parameter
        self.WATCH_FOLDER = watch_folder
        self.API_URL = "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation"
        self.SUPPORTED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.webp'}

    def call_qwen_vl(self, image_path):
        """调用千问VL API识别图片内容"""
        try:
            # 检查文件是否完整（防止文件正在写入）
            if not self._is_file_ready(image_path):
                print(f"⏳ 文件可能正在写入，稍后重试: {os.path.basename(image_path)}")
                return False

            with open(image_path, "rb") as f:
                img_b64 = base64.b64encode(f.read()).decode()

            headers = {
                "Authorization": f"Bearer {self.API_KEY}",
                "Content-Type": "application/json"
            }

            payload = {
                "model": "qwen-vl-plus",
                "input": {
                    "messages": [
                        {
                            "role": "user",
                            "content": [
                                {
                                    "image": f"data:image/jpeg;base64,{img_b64}"
                                },
                                {
                                    "text": "材质  规格   公司  厚度  产地  状态  价格"
                                }
                            ]
                        }
                    ]
                },
                "parameters": {
                    "temperature": 0.3
                }
            }

            print(f"🔄 正在识别: {os.path.basename(image_path)}")
            resp = requests.post(self.API_URL, headers=headers, json=payload, timeout=30)

            if resp.status_code == 200:
                j = resp.json()
                try:
                    result = j["output"]["results"][0]["output_text"]
                except (KeyError, IndexError) as e:
                    result = f"解析响应时出错: {e}\n原始响应: {j}"

                print(f"\n✅ {os.path.basename(image_path)} 识别完成！")
                self._save_result(image_path, result)
                return True
            else:
                print(f"❌ API请求失败 {resp.status_code} - {resp.text}")
                return False

        except Exception as e:
            print(f"❌ 处理图片时出错: {e}")
            return False

    def _is_file_ready(self, file_path):
        """检查文件是否准备好（非正在写入状态）"""
        try:
            # 尝试以追加模式打开，如果文件正在被其他进程写入，可能会失败
            with open(file_path, 'a'):
                pass
            return True
        except IOError:
            return False

    def _save_result(self, image_path, result):
        try:
            txt_path = "specs.txt"  # 统一保存为 specs.txt
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(result)
            print(f"💾 结果已保存至: {txt_path}")
        except Exception as e:
            print(f"❌ 保存结果失败: {e}")


class ImageHandler(FileSystemEventHandler):
    def __init__(self, processor):
        self.processor = processor
        self.processed_files = set()  # 防止重复处理

    def on_created(self, event):
        """处理文件创建事件"""
        if not event.is_directory:
            file_path = event.src_path
            _, ext = os.path.splitext(file_path)

            if ext.lower() in self.processor.SUPPORTED_EXTENSIONS:
                # 防止重复处理
                if file_path in self.processed_files:
                    return
                self.processed_files.add(file_path)

                print(f"\n📸 检测到新图片: {os.path.basename(file_path)}")
                # 等待文件完全写入
                time.sleep(2)
                self.processor.call_qwen_vl(file_path)


def main():
    # 询问用户输入API密钥
    API_KEY = input("请输入您的API密钥: ").strip()
    WATCH_FOLDER = r"D:\projects"

    # 检查监控文件夹是否存在
    if not os.path.exists(WATCH_FOLDER):
        print(f"❌ 监控文件夹不存在: {WATCH_FOLDER}")
        return

    print(f"👀 正在监听文件夹: {WATCH_FOLDER}")
    print("📁 支持的文件格式: jpg, jpeg, png, bmp, webp")
    print("⏹️ 按 Ctrl+C 停止监控")

    # 事件处理器
    processor = QwenVLProcessor(API_KEY, WATCH_FOLDER)
    handler = ImageHandler(processor)
    observer = Observer()
    observer.schedule(handler, WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(2)
    except KeyboardInterrupt:
        print("\n👋 停止监控...")
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
