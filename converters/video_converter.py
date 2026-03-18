"""
视频转文字转换器
注意：
1. 如果遇到 "No module named 'moviepy.editor'" 错误，可能是 moviepy 版本问题。
   解决方案：
   - 卸载当前版本：pip uninstall moviepy
   - 安装指定版本：pip install moviepy==1.0.3

2. 如果遇到 ffmpeg 相关错误，需要正确设置 ffmpeg 路径。
   解决方案：
   - 确保已安装 ffmpeg
   - 设置 FFMPEG_PATH 为 ffmpeg 可执行文件的完整路径
   - 如果使用 winget 安装的 ffmpeg，路径通常在：
     C:/Users/[用户名]/AppData/Local/Microsoft/WinGet/Packages/Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe/ffmpeg-7.1.1-full_build/bin/ffmpeg.exe
   - 如果使用其他方式安装，请相应修改路径

3. 如果遇到 "No module named 'opencc'" 错误，需要正确安装 OpenCC。
   解决方案：
   - 卸载当前版本：pip uninstall opencc-python-reimplemented
   - 安装指定版本：pip install opencc-python-reimplemented==0.1.7
   - 注意：不要使用 opencc-python 包，而是使用 opencc-python-reimplemented

4. 如果遇到其他依赖包缺失错误，需要安装以下包：
   解决方案：
   - pip install whisper  # 用于语音识别
   - pip install torch  # PyTorch 深度学习框架
   - pip install numpy  # 数值计算库
   - pip install opencc-python-reimplemented==0.1.7  # 用于繁简转换
   - pip install moviepy==1.0.3  # 用于视频处理
"""

import os
import sys
import json
import subprocess
import tempfile
import time
from moviepy.editor import VideoFileClip
import whisper
from opencc import OpenCC
import logging

# 设置 ffmpeg 路径
FFMPEG_PATH = r"C:\Users\Grant\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-7.1.1-full_build\bin\ffmpeg.exe"
os.environ["IMAGEIO_FFMPEG_EXE"] = FFMPEG_PATH
os.environ["PATH"] = os.path.dirname(FFMPEG_PATH) + os.pathsep + os.environ["PATH"]

# 设置标准输出和错误流的编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# 重定向 MoviePy 的输出到 stderr
logging.getLogger('moviepy').setLevel(logging.ERROR)

# 发送错误消息（可选）
def send_error(error):
    """发送错误消息，使用 content 替代 error"""
    print(json.dumps({
        "type": "error",
        "content": error       # 错误信息
    }, ensure_ascii=False), flush=True)

# 发送进度信息
def send_progress(progress, content="处理中"):
    """发送进度更新，统一使用 content 字段"""
    print(json.dumps({
        "type": "progress",
        "progress": progress,  # 进度值，0-100
        "content": content     # 描述性内容
    }, ensure_ascii=False), flush=True)

# 发送信息消息，不刷新输出框，只在输出框追加
def send_info(content):
    """发送信息消息，统一使用 content 字段"""
    print(json.dumps({
        "type": "info",
        "content": content,     # 主要输出内容
        "append": True         # 是否追加，保留原有逻辑
    }, ensure_ascii=False), flush=True)

# 发送完成消息，刷新输出框，并把结果文本输出到输出框
def send_complete(content):
    """发送完成消息，统一使用 content 字段"""
    print(json.dumps({
        "type": "complete",
        "content": content,    # 完成后的结果
        "append": False         # 是否追加，保留原有逻辑
    }, ensure_ascii=False), flush=True)

# 清空输出框
def clear_output():
    """清空输出框"""
    print(json.dumps({
        "type": "complete",
        "content": "",
        "append": False
    }, ensure_ascii=False), flush=True)

def convert_video_to_markdown(video_path):
    """Convert video to markdown using Whisper"""
    model = None
    video = None
    temp_audio_path = None
    
    try:
        # 检查视频文件是否存在
        if not os.path.exists(video_path):
            send_error(f"视频文件不存在: {video_path}")
            return

        # 检查视频文件是否可读
        if not os.access(video_path, os.R_OK):
            send_error(f"视频文件不可读: {video_path}")
            return

        # 发送进度更新
        send_progress(10, "正在加载Whisper模型...")
        
        # 加载Whisper模型
        try:
            model = whisper.load_model("base")
            send_progress(20, "Whisper模型加载成功")
        except Exception as e:
            send_error(f"加载Whisper模型失败: {str(e)}")
            return

        # 创建临时音频文件
        temp_audio = f"temp_audio_{os.getpid()}_{int(time.time())}.mp3"
        temp_audio_path = os.path.join(os.path.dirname(video_path), temp_audio)
        
        send_progress(30, "正在从视频中提取音频...")
        
        # 提取音频
        try:
            video = VideoFileClip(video_path)
            video.audio.write_audiofile(temp_audio_path)
            video.close()
            video = None  # 确保视频对象被释放
            send_progress(40, "音频提取完成")
        except Exception as e:
            send_error(f"提取音频失败: {str(e)}")
            return

        # 检查音频文件是否成功创建
        if not os.path.exists(temp_audio_path):
            send_error("音频文件创建失败")
            return

        # 检查音频文件是否可读
        if not os.access(temp_audio_path, os.R_OK):
            send_error("音频文件不可读")
            return

        # 检查音频文件大小
        if os.path.getsize(temp_audio_path) == 0:
            send_error("音频文件为空")
            return

        send_progress(50, "正在进行语音识别...")
        
        # 使用Whisper进行语音识别
        try:
            # 使用绝对路径而不是切换目录
            send_info(f"开始语音识别，音频文件路径: {temp_audio_path}")
            
            # 检查音频文件大小
            audio_size = os.path.getsize(temp_audio_path)
            send_info(f"音频文件大小: {audio_size / 1024 / 1024:.2f} MB")
            
            # 设置 Whisper 的转录参数
            result = model.transcribe(
                temp_audio_path,
                language="zh",  # 指定语言为中文
                task="transcribe",  # 指定任务为转录
                fp16=False,  # 使用 FP32 以提高稳定性
                beam_size=5,  # 增加 beam size 以提高准确性
                best_of=5,  # 增加候选数量
                temperature=0.0,  # 使用确定性采样
                condition_on_previous_text=True,  # 考虑上下文
                no_speech_threshold=0.6,  # 调整无语音阈值
                logprob_threshold=-1.0,  # 调整日志概率阈值
                compression_ratio_threshold=1.2  # 调整压缩比阈值
            )
            
            # 检查转录结果
            if not result or 'segments' not in result:
                send_info("警告：转录结果为空或格式不正确")
                send_error("语音识别结果为空")
                return
                
            # 检查片段数量
            segments_count = len(result['segments'])
            send_info(f"语音识别完成，识别到 {segments_count} 个片段")
            
            # 检查每个片段的内容
            valid_segments = 0
            total_text_length = 0
            for i, segment in enumerate(result['segments']):
                if segment.get('text', '').strip():
                    valid_segments += 1
                    total_text_length += len(segment['text'])
                    if i < 3:  # 只显示前3个片段
                        send_info(f"片段 {i+1}: {segment['text'][:50]}...")
            
            # 打印统计信息
            send_info(f"有效片段数: {valid_segments}/{segments_count}, 总文本长度: {total_text_length} 字符")
            
            if valid_segments == 0:
                send_info("警告：没有识别到任何有效文本")
                send_error("语音识别结果为空")
                return
            
            send_progress(80, "语音识别完成")
        except Exception as e:
            send_info(f"语音识别过程出错: {str(e)}")
            send_error(f"语音识别失败: {str(e)}")
            return

        # 将繁体中文转换为简体中文
        try:
            send_progress(85, "正在进行繁简转换...")
            cc = OpenCC('t2s')
            
            # 处理每个片段，保留时间戳
            processed_segments = []
            for segment in result["segments"]:
                start_time = int(segment["start"])
                minutes = start_time // 60
                seconds = start_time % 60
                timestamp = f"{minutes:02d}:{seconds:02d}"
                
                # 只转换文本内容，保留时间戳
                text = segment["text"].strip()
                if text:
                    simplified_text = cc.convert(text)
                    processed_segments.append(f"**{timestamp}** {simplified_text}\n\n")
            
            # 检查是否有处理后的文本
            if not processed_segments:
                send_error("语音识别结果为空")
                return
                
            # 将所有片段组合成完整的markdown文本
            markdown = "# 视频转文字\n\n" + "\n\n".join(processed_segments)
            send_progress(90, "繁简转换完成")
            
            # 清空输出框并发送最终结果
            clear_output()  # 先清空输出框
            print(markdown)  # 直接打印结果文本
            
        except Exception as e:
            send_error(f"繁简转换失败: {str(e)}")
            return

        send_progress(90, "正在清理临时文件...")
        
        # 清理临时文件
        try:
            if os.path.exists(temp_audio_path):
                os.remove(temp_audio_path)
            send_progress(100, "转换完成")
        except Exception as e:
            # 如果删除失败，记录错误但不影响结果
            send_info(f"保留临时音频文件: {temp_audio_path}")
            
            # 清空输出框并发送最终结果
            clear_output()  # 先清空输出框
            print(markdown)  # 直接打印结果文本

    except Exception as e:
        send_error(f"转换失败: {str(e)}")
        return
    finally:
        # 清理资源
        try:
            if video:
                video.close()
            if model:
                del model
            if temp_audio_path and os.path.exists(temp_audio_path):
                try:
                    os.remove(temp_audio_path)
                except:
                    pass
            # 强制进行垃圾回收
            import gc
            gc.collect()
        except Exception as e:
            send_info(f"清理资源时出错: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        send_error("请提供视频文件路径")
        sys.exit(1)
    
    video_path = sys.argv[1]
    convert_video_to_markdown(video_path) 