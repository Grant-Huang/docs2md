/**
 * 文档资料知识化APP 前端逻辑
 */
(function () {
    const fileInput = document.getElementById("fileInput");
    const dirInput = document.getElementById("dirInput");
    const btnSelectFile = document.getElementById("btnSelectFile");
    const btnSelectDir = document.getElementById("btnSelectDir");
    const inputHint = document.getElementById("inputHint");

    btnSelectFile.addEventListener("click", () => {
        dirInput.value = "";
        fileInput.click();
    });
    btnSelectDir.addEventListener("click", () => {
        fileInput.value = "";
        dirInput.click();
    });

    function updateInputHint() {
        const files = getSupportedFiles();
        if (files.length === 0) {
            inputHint.textContent = "未选择";
        } else if (files.length === 1) {
            inputHint.textContent = "已选: " + files[0].name;
        } else {
            inputHint.textContent = "已选 " + files.length + " 个文件";
        }
    }

    fileInput.addEventListener("change", () => {
        if (fileInput.files && fileInput.files.length > 0) dirInput.value = "";
        updateInputHint();
    });
    dirInput.addEventListener("change", () => {
        if (dirInput.files && dirInput.files.length > 0) fileInput.value = "";
        updateInputHint();
    });
    const btnSelectOutputDir = document.getElementById("btnSelectOutputDir");
    const outputDirHint = document.getElementById("outputDirHint");
    const outputDirFallback = document.getElementById("outputDirFallback");
    const outputDirSelect = document.getElementById("outputDirSelect");

    let outputDirHandle = null;

    const hasDirPicker = "showDirectoryPicker" in window;

    if (!hasDirPicker) {
        outputDirFallback.style.display = "block";
        outputDirHint.textContent = "使用下方服务器子目录";
    }

    btnSelectOutputDir.addEventListener("click", async () => {
        if (!hasDirPicker) return;
        try {
            outputDirHandle = await window.showDirectoryPicker({ mode: "readwrite" });
            outputDirHint.textContent = "已选: " + outputDirHandle.name;
        } catch (e) {
            if (e.name !== "AbortError") {
                outputDirHint.textContent = "选择失败: " + e.message;
            }
        }
    });

    function getOutputDir() {
        if (outputDirHandle) return "";
        if (outputDirSelect) return outputDirSelect.value || "";
        return "";
    }

    async function writeToOutputDir(relativePath, content) {
        if (!outputDirHandle) return;
        const parts = relativePath.replace(/\\/g, "/").split("/");
        let dir = outputDirHandle;
        for (let i = 0; i < parts.length - 1; i++) {
            dir = await dir.getDirectoryHandle(parts[i], { create: true });
        }
        const file = await dir.getFileHandle(parts[parts.length - 1], { create: true });
        const w = await file.createWritable();
        await w.write(typeof content === "string" ? content : content);
        await w.close();
    }
    const format = document.getElementById("format");
    const convertBtn = document.getElementById("convertBtn");
    const progLog = document.getElementById("progLog");
    const resultSingle = document.getElementById("resultSingle");
    const resultRendered = document.getElementById("resultRendered");
    const resultRaw = document.getElementById("resultRaw");
    const resultIndex = document.getElementById("resultIndex");
    const indexLinks = document.getElementById("indexLinks");
    const resultEmpty = document.getElementById("resultEmpty");

    function log(msg) {
        const cur = progLog.textContent || "";
        progLog.textContent = cur ? cur + "\n" + msg : msg;
        progLog.scrollTop = progLog.scrollHeight;
    }

    function setResult(content, isMarkdown) {
        resultSingle.style.display = "block";
        resultIndex.style.display = "none";
        resultEmpty.style.display = "none";
        if (isMarkdown && typeof marked !== "undefined") {
            const md = (content || "").replace(
                /\]\(assets\//g,
                "](/output/assets/"
            );
            resultRendered.innerHTML = marked.parse(md);
            resultRendered.style.display = "block";
            resultRaw.style.display = "none";
        } else {
            resultRaw.textContent = content || "";
            resultRaw.style.display = "block";
            resultRendered.style.display = "none";
        }
    }

    function setIndex(results) {
        resultSingle.style.display = "none";
        resultIndex.style.display = "block";
        resultEmpty.style.display = "none";
        indexLinks.innerHTML = "";
        const items = results.filter((r) => r.output && !r.path.includes("index"));
        items.forEach((r) => {
            const a = document.createElement("a");
            const relPath = r.output.replace(/^.*[\\/]output[\\/]?/, "").replace(/\\/g, "/");
            a.href = "/output/" + relPath;
            a.textContent = r.path.split(/[/\\]/).pop() || r.output;
            a.target = "_blank";
            indexLinks.appendChild(a);
        });
        const indexItem = results.find((r) => r.path === "index");
        if (indexItem && indexItem.output) {
            const a = document.createElement("a");
            a.href = "/output/index.md";
            a.textContent = "index.md";
            a.target = "_blank";
            a.className = "fw-bold";
            indexLinks.insertBefore(a, indexLinks.firstChild);
        }
    }

    function resetResult() {
        resultSingle.style.display = "none";
        resultIndex.style.display = "none";
        resultEmpty.style.display = "block";
    }

    async function convertSingleFile(file) {
        const formData = new FormData();
        formData.append("file", file);
        formData.append("output_dir", getOutputDir());
        formData.append("format", format.value);

        log("开始转换: " + file.name);
        const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
        log("正在上传文件（" + sizeMB + " MB），请稍候，大文件可能需要数分钟...");
        const resp = await fetch("/api/convert", {
            method: "POST",
            body: formData,
        });

        if (!resp.ok) {
            log("请求失败: " + resp.status);
            return;
        }

        const reader = resp.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";

        while (true) {
            const { value, done } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const parts = buffer.split(/\n\ndata: /);
            buffer = parts.pop() || "";
            for (let i = 0; i < parts.length; i++) {
                const p = i === 0 ? parts[i] : "data: " + parts[i];
                const jsonStr = p.startsWith("data: ") ? p.slice(6).trim() : p.trim();
                if (!jsonStr) continue;
                try {
                    const data = JSON.parse(jsonStr);
                    if (data.type === "debug") log(data.content);
                    else if (data.type === "partial") {
                        setResult(data.content, format.value === "md");
                    } else if (data.type === "error") {
                        log("错误: " + data.content);
                        return;
                    } else if (data.type === "complete") {
                        log("转换完成");
                        setResult(data.content, format.value === "md");
                        if (outputDirHandle && data.content) {
                            const ext = format.value === "md" ? ".md" : ".txt";
                            const base = file.name.replace(/\.[^.]+$/, "");
                            const outName = base + ext;
                            try {
                                await writeToOutputDir(outName, data.content);
                                log("已保存到选中文件夹: " + outName);
                            } catch (e) {
                                log("保存到本地失败: " + e.message);
                            }
                        }
                    }
                } catch (e) {}
            }
        }
        if (buffer.trim()) {
            try {
                const jsonStr = buffer.replace(/^data: /, "").trim();
                const data = JSON.parse(jsonStr);
                if (data.type === "debug") log(data.content);
                else if (data.type === "partial") {
                    setResult(data.content, format.value === "md");
                } else if (data.type === "complete") {
                    setResult(data.content, format.value === "md");
                    if (outputDirHandle && data.content) {
                        const ext = format.value === "md" ? ".md" : ".txt";
                        const base = file.name.replace(/\.[^.]+$/, "");
                        const outName = base + ext;
                        try {
                            await writeToOutputDir(outName, data.content);
                            log("已保存到选中文件夹: " + outName);
                        } catch (e) {
                            log("保存到本地失败: " + e.message);
                        }
                    }
                }
            } catch (e) {}
        }
    }

    async function convertDirectory(files) {
        const formData = new FormData();
        for (const f of files) {
            if (/\.(docx?|xlsx?)$/i.test(f.name)) {
                formData.append("files", f, f.webkitRelativePath || f.name);
            }
        }
        formData.append("output_dir", getOutputDir());
        formData.append("format", format.value);

        const fileCount = formData.getAll("files").length;
        log("正在上传 " + fileCount + " 个文件，请稍候...");
        const resp = await fetch("/api/convert-dir-upload", {
            method: "POST",
            body: formData,
        });

        if (!resp.ok) {
            try {
                const err = await resp.json();
                log("错误: " + (err.detail || resp.statusText));
            } catch (_) {
                log("请求失败: " + resp.status);
            }
            return;
        }

        // SSE 流式读取进度
        const reader = resp.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";

        async function handleComplete(results) {
            log("批量转换完成");
            if (results && results.length) {
                setIndex(results);
                if (outputDirHandle) {
                    const items = results.filter(
                        (r) => r.output && !r.path.includes("index")
                    );
                    for (const r of items) {
                        const relPath = r.output
                            .replace(/^.*[\\/]output[\\/]?/, "")
                            .replace(/\\/g, "/");
                        if (!relPath) continue;
                        try {
                            const fr = await fetch("/output/" + relPath);
                            if (fr.ok) {
                                const blob = await fr.blob();
                                await writeToOutputDir(relPath, blob);
                                log("已保存: " + relPath);
                            }
                        } catch (e) {
                            log("保存失败 " + relPath + ": " + e.message);
                        }
                    }
                    const idxItem = results.find((r) => r.path === "index");
                    if (idxItem && idxItem.content) {
                        try {
                            await writeToOutputDir("index.md", idxItem.content);
                            log("已保存: index.md");
                        } catch (e) {
                            log("保存 index 失败: " + e.message);
                        }
                    }
                }
            } else {
                setResult("无有效文件", false);
            }
        }

        function processChunk(chunk) {
            const parts = chunk.split(/\n\ndata: /);
            const last = parts.pop() || "";
            for (let i = 0; i < parts.length; i++) {
                const p = i === 0 ? parts[i] : "data: " + parts[i];
                const jsonStr = p.startsWith("data: ") ? p.slice(6).trim() : p.trim();
                if (!jsonStr) continue;
                try {
                    const data = JSON.parse(jsonStr);
                    if (data.type === "debug") log(data.content);
                    else if (data.type === "error") log("错误: " + data.content);
                    else if (data.type === "complete") handleComplete(data.results);
                } catch (_) {}
            }
            return last;
        }

        while (true) {
            const { value, done } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            buffer = processChunk(buffer);
        }
        if (buffer.trim()) processChunk(buffer + "\n\n");
    }

    function getSupportedFiles() {
        const fromFile = Array.from(fileInput.files || []);
        const fromDir = Array.from(dirInput.files || []);
        const all = fromFile.length > 0 ? fromFile : fromDir;
        return all.filter((f) => /\.(docx?|xlsx?)$/i.test(f.name));
    }

    convertBtn.addEventListener("click", async () => {
        progLog.textContent = "";
        resetResult();
        convertBtn.disabled = true;

        try {
            const files = getSupportedFiles();
            if (files.length === 0) {
                const hasAny = (fileInput.files && fileInput.files.length > 0) ||
                    (dirInput.files && dirInput.files.length > 0);
                if (hasAny) {
                    log("未找到支持的文档 (.docx/.xlsx)");
                } else {
                    log("请选择文件或文件夹");
                }
            } else if (files.length === 1) {
                await convertSingleFile(files[0]);
            } else {
                await convertDirectory(files);
            }
        } catch (e) {
            log("错误: " + e.message);
        } finally {
            convertBtn.disabled = false;
        }
    });
})();
