const express = require('express');
const multer = require('multer');
const path = require('path');
const { spawn } = require('child_process');
const cors = require('cors');
const fs = require('fs');
const { spawn: spawnPython } = require('child_process');

const app = express();
const port = 3000;

// 启用CORS
app.use(cors());

// 添加 JSON 解析中间件
app.use(express.json());

// 添加 CSP 头
app.use((req, res, next) => {
    res.setHeader(
      'Content-Security-Policy',
      "default-src 'self'; " +
      "script-src 'self' 'unsafe-inline' 'unsafe-eval' https://cdn.jsdelivr.net https://cdnjs.cloudflare.com; " +
      "style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; " +
      "font-src 'self' https://cdn.jsdelivr.net; " +
      "img-src 'self' data: blob:; " +
      "media-src 'self' data: blob:; " +
      "connect-src 'self'"
    );
    next();
  });

// 提供静态文件服务
app.use(express.static('public'));
app.use('/uploads', express.static('uploads'));

// 确保必要的目录存在
['uploads', 'public'].forEach(dir => {
    if (!fs.existsSync(dir)) {
        console.log(`创建目录: ${dir}`);
        fs.mkdirSync(dir);
    } else {
        //console.log(`目录已存在: ${dir}`);
    }
});

// 配置文件上传
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const uploadDir = path.join(__dirname, 'uploads');
        console.log(`检查上传目录: ${uploadDir}`);
        if (!fs.existsSync(uploadDir)) {
            console.log(`创建上传目录: ${uploadDir}`);
            fs.mkdirSync(uploadDir, { recursive: true });
        }
        console.log(`上传目录状态: ${fs.existsSync(uploadDir) ? '存在' : '不存在'}`);
        cb(null, uploadDir);
    },
    filename: function (req, file, cb) {
        const filename = Date.now() + path.extname(file.originalname);
        console.log(`生成文件名: ${filename}`);
        cb(null, filename);
    }
});

const upload = multer({ 
    storage: storage,
    limits: {
        fileSize: 100 * 1024 * 1024, // 限制文件大小为100MB
    },
    fileFilter: function (req, file, cb) {
        // 检查文件类型
        const allowedTypes = [
            'application/pdf',
            'video/mp4', 'video/quicktime', 'video/x-msvideo',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/msword',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        if (!allowedTypes.includes(file.mimetype)) {
            console.error('不支持的文件类型:', file.mimetype);
            return cb(new Error('不支持的文件类型'));
        }
        cb(null, true);
    }
});

// 处理文件上传和转换
app.post('/convert', upload.single('file'), (req, res) => {
    if (!req.file) {
        console.error('没有上传文件');
        return res.status(400).json({ error: '没有上传文件' });
    }

    const filePath = req.file.path;
    const format = req.body.format || 'text';
    const fileType = req.file.mimetype;

    console.log('文件上传信息:', {
        originalname: req.file.originalname,
        filename: req.file.filename,
        path: req.file.path,
        size: req.file.size,
        mimetype: req.file.mimetype
    });

    // 检查文件是否存在
    if (!fs.existsSync(filePath)) {
        console.error('文件不存在:', filePath);
        return res.status(400).json({ error: '文件不存在' });
    }

    // 检查文件大小
    const stats = fs.statSync(filePath);
    console.log('文件状态:', {
        size: stats.size,
        isFile: stats.isFile(),
        isDirectory: stats.isDirectory()
    });

    if (stats.size === 0) {
        console.error('文件为空:', filePath);
        fs.unlinkSync(filePath);
        return res.status(400).json({ error: '文件为空' });
    }

    // 根据文件类型选择转换脚本
    let pythonScript;
    if (fileType === 'application/pdf') {
        pythonScript = path.join(__dirname, 'converters', 'pdf_converter.py');
    } else if (fileType.startsWith('video/')) {
        pythonScript = path.join(__dirname, 'converters', 'video_converter.py');
    } else if (fileType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
               fileType === 'application/msword') {
        pythonScript = path.join(__dirname, 'converters', 'doc_converter.py');
    } else if (fileType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
               fileType === 'application/vnd.ms-excel') {
        pythonScript = path.join(__dirname, 'converters', 'excel_converter.py');
    } else {
        return res.status(400).json({ error: '不支持的文件类型' });
    }

    // 设置响应头
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    // 调用Python脚本进行转换
    const pythonProcess = spawn(path.join(__dirname, '.venv', 'Scripts', 'python.exe'), [pythonScript, filePath, format], {
        stdio: ['pipe', 'pipe', 'pipe'],
        encoding: 'utf8',
        shell: false,  // Don't use shell to avoid buffering issues
        windowsHide: true,
        env: {
            ...process.env,
            PYTHONPATH: process.env.PYTHONPATH || '',
            PYTHONIOENCODING: 'utf-8',
            PYTHONUNBUFFERED: '1'
        }
    });

    let stdoutBuffer = '';
    let errorOutput = '';
    let hasCompleted = false;

    function processSSEMessages(buffer) {
        const parts = buffer.split(/\n\ndata: /);
        const messages = parts.map((p, i) => (i === 0 ? p : 'data: ' + p));
        let remainder = '';
        if (messages.length > 0) {
            const last = messages[messages.length - 1];
            try {
                if (last.startsWith('data: ')) {
                    JSON.parse(last.slice(6).trim());
                }
            } catch (e) {
                remainder = messages.pop() || '';
            }
        }
        for (const msg of messages) {
            if (msg.startsWith('data: ')) {
                try {
                    const data = JSON.parse(msg.slice(6).trim());
                    if (data.type === 'progress') {
                        res.write(`data: ${JSON.stringify(data)}\n\n`);
                    } else if (data.type === 'debug') {
                        res.write(`data: ${JSON.stringify(data)}\n\n`);
                    } else if (data.type === 'complete') {
                        hasCompleted = true;
                        res.write(`data: ${JSON.stringify(data)}\n\n`);
                    } else if (data.type === 'error') {
                        errorOutput += (data.content || data.message || '') + '\n';
                        res.write(`data: ${JSON.stringify({ type: 'error', content: data.content || data.message })}\n\n`);
                    }
                } catch (e) {
                    console.error('解析 Python 输出失败:', msg, e);
                }
            }
        }
        return remainder;
    }

    pythonProcess.stdout.on('data', (data) => {
        stdoutBuffer += data.toString();
        stdoutBuffer = processSSEMessages(stdoutBuffer);
    });

    pythonProcess.stderr.on('data', (data) => {
        console.error('Python错误:', data.toString());
        errorOutput += data.toString();
    });

    pythonProcess.on('error', (err) => {
        console.error('Python进程错误:', err);
        console.error('错误详情:', {
            message: err.message,
            code: err.code,
            stack: err.stack
        });
        errorOutput += err.toString();
    });

    pythonProcess.on('close', (code) => {
        console.log('Python进程退出，代码:', code);

        // 处理剩余缓冲区
        if (stdoutBuffer.trim()) {
            processSSEMessages(stdoutBuffer + '\n\n');
        }

        // 删除上传的文件
        fs.unlink(filePath, (err) => {
            if (err) {
                console.error('删除文件失败:', err);
            } else {
                console.log('文件已删除:', filePath);
            }
        });

        if (!hasCompleted) {
            if (code !== 0 || errorOutput) {
                res.write(`data: ${JSON.stringify({ type: 'error', content: errorOutput || '转换失败' })}\n\n`);
            } else {
                res.write(`data: ${JSON.stringify({ type: 'error', content: '未收到转换结果' })}\n\n`);
            }
        }
        res.end();
    });
});

// YouTube转换路由
app.post('/api/tools/youtube-to-markdown', async (req, res) => {
    try {
        if (!req.body.url) {
            return res.status(400).json({ error: '缺少 url 参数' });
        }

        const url = req.body.url;
        
        // 从URL中提取视频ID
        let videoId;
        if (url.includes('youtube.com')) {
            videoId = url.split('v=')[1].split('&')[0];
        } else if (url.includes('youtu.be')) {
            videoId = url.split('/')[-1];
        } else {
            return res.status(400).json({ error: '无效的YouTube URL' });
        }

        // 调用Python脚本进行转换
        const pythonScript = path.join(__dirname, 'converters', 'youtube_converter.py');
        const pythonProcess = spawnPython('D:\\tools\\python3\\python.exe', [pythonScript, videoId]);

        let output = '';
        let errorOutput = '';

        pythonProcess.stdout.on('data', (data) => {
            console.log('Python输出:', data.toString());
            output += data.toString();
        });

        pythonProcess.stderr.on('data', (data) => {
            console.error('Python错误:', data.toString());
            errorOutput += data.toString();
        });

        pythonProcess.on('error', (err) => {
            console.error('Python进程错误:', err);
            console.error('错误详情:', {
                message: err.message,
                code: err.code,
                stack: err.stack
            });
            errorOutput += err.toString();
        });

        pythonProcess.on('close', (code) => {
            console.log('Python进程退出，代码:', code);
            
            if (code === 0 && output) {
                res.json({ text: output });
            } else {
                res.status(500).json({ 
                    error: '转换失败',
                    details: errorOutput || '未知错误'
                });
            }
        });
    } catch (err) {
        console.error('服务器错误:', err);
        res.status(500).json({ error: err.message || '服务器内部错误' });
    }
});

// Video conversion route
app.post('/convert-video', upload.single('video'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No video file uploaded' });
        }

        const videoPath = path.resolve(req.file.path);
        console.log('Processing video file:', {
            originalPath: req.file.path,
            resolvedPath: videoPath,
            exists: fs.existsSync(videoPath),
            isFile: fs.statSync(videoPath).isFile(),
            size: fs.statSync(videoPath).size
        });
        
        // 检查文件是否存在
        if (!fs.existsSync(videoPath)) {
            console.error('Uploaded file not found:', videoPath);
            return res.status(400).json({ error: 'Uploaded file not found' });
        }

        // 检查文件是否可读
        try {
            fs.accessSync(videoPath, fs.constants.R_OK);
        } catch (err) {
            console.error('File is not readable:', videoPath);
            return res.status(400).json({ error: 'File is not readable' });
        }

        // 设置响应头，启用 Server-Sent Events (SSE)
        res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');

        const pythonScript = path.resolve(__dirname, 'converters', 'video_converter.py');
        
        // 检查Python脚本是否存在
        if (!fs.existsSync(pythonScript)) {
            console.error('Python script not found:', pythonScript);
            res.write(`data: ${JSON.stringify({ type: 'error', error: 'Python script not found' })}\n\n`);
            res.end();
            return;
        }

        console.log('Starting Python process with:', {
            script: pythonScript,
            videoPath: videoPath,
            scriptExists: fs.existsSync(pythonScript),
            videoExists: fs.existsSync(videoPath)
        });

        // 启动Python进程进行视频转换
        const pythonProcess = spawn(path.join(__dirname, '.venv', 'Scripts', 'python.exe'), [
            pythonScript,
            videoPath
        ], {
            stdio: ['pipe', 'pipe', 'pipe'],
            encoding: 'utf8',
            shell: false,
            windowsHide: true,
            env: {
                ...process.env,
                PYTHONPATH: process.env.PYTHONPATH || '',
                PYTHONIOENCODING: 'utf-8'
            }
        });

        let output = '';
        let errorOutput = '';

        // 处理Python进程的输出
        pythonProcess.stdout.on('data', (data) => {
            const outputStr = data.toString().trim();
            console.log('Received from Python:', outputStr);
            try {
                // 尝试解析JSON输出
                const jsonData = JSON.parse(outputStr);
                if (jsonData.type === 'progress') {
                    // 进度信息只发送到进度条
                    res.write(`data: ${outputStr}\n\n`);
                } else if (jsonData.type === 'info') {
                    // 音频转换文本发送到输出框
                    res.write(`data: ${outputStr}\n\n`);
                }
                else if (jsonData.type === 'complete') {
                    // 音频转换文本发送到输出框
                    res.write(`data: ${outputStr}\n\n`);
                }
            } catch (e) {
                res.write(`data: ${outputStr}\n\n`);
                console.warn('Ignoring non-JSON output:', outputStr);
            }
        });

        // 处理Python进程的错误输出
        pythonProcess.stderr.on('data', (data) => {
            console.error('Python error:', data.toString());
            errorOutput += data.toString();
        });

        // 处理Python进程的错误
        pythonProcess.on('error', (err) => {
            console.error('Python process error:', err);
            errorOutput += err.toString();
        });

        // 处理Python进程的结束
        pythonProcess.on('close', (code) => {
            console.log('Python process closed with code:', code);
            
            // 清理上传的文件
            try {
                if (fs.existsSync(videoPath)) {
                    fs.unlinkSync(videoPath);
                    console.log('Cleaned up uploaded file:', videoPath);
                }
            } catch (err) {
                console.error('Failed to clean up file:', err);
            }

            if (code !== 0) {
                // 如果进程异常退出，发送错误信息
                res.write(`data: ${JSON.stringify({ type: 'error', error: errorOutput || 'Unknown error' })}\n\n`);
            }
            res.end();
        });
    } catch (error) {
        console.error('Error converting video:', error);
        res.write(`data: ${JSON.stringify({ type: 'error', error: error.message })}\n\n`);
        res.end();
    }
});

// 错误处理中间件
app.use((err, req, res, next) => {
    console.error('服务器错误:', err);
    if (err instanceof multer.MulterError) {
        if (err.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({ error: '文件大小超过限制（最大100MB）' });
        }
    }
    res.status(500).json({ error: err.message || '服务器内部错误' });
});

// 添加健康检查端点
app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

// 启动服务器
app.listen(port, () => {
  console.log(`服务器运行在 http://localhost:${port}`);
}); 