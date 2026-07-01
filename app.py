"""
闪电仓计算工具 - Web 版
基于现有 lightning_warehouse_tool.py 的核心逻辑，包装成 Flask Web 服务。

启动:  python app.py
访问:  http://localhost:5000
"""
import os
import re
import uuid
import shutil
import tempfile
import logging
import time
import json
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from flask import Flask, request, render_template_string, send_file, jsonify, redirect, url_for

# 复用现有工具的核心函数（直接 import 闪电仓工具模块）
# 关键函数: extract_short_name, make_pivot_file2, process_file1
# 为了避免路径问题，这里把核心逻辑重新写一遍（精简版）
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from lightning_warehouse_tool import (
        extract_short_name,
        make_pivot_file2,
        make_pivot_file3,
        make_pivot_file4,
        process_file1,
    )
except ImportError as e:
    # 如果是 PyQt5 缺失错误（云端环境），那是因为 module 顶层 import 了 PyQt5
    # 模块代码已用 try/except 处理，这里再兜底
    if 'PyQt5' not in str(e):
        raise  # 其他 ImportError 才是真问题，重抛
    # 重新尝试，云端环境下 PyQt5 是可选的
    import importlib
    import lightning_warehouse_tool
    extract_short_name = lightning_warehouse_tool.extract_short_name
    make_pivot_file2 = lightning_warehouse_tool.make_pivot_file2
    make_pivot_file3 = lightning_warehouse_tool.make_pivot_file3
    make_pivot_file4 = lightning_warehouse_tool.make_pivot_file4
    process_file1 = lightning_warehouse_tool.process_file1

# ============== Flask 应用 ==============
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 上限

# 任务存储（用文件存储，以适配 gunicorn 多 worker 场景）
TASK_DIR = os.path.join(tempfile.gettempdir(), 'lightning_tasks')
os.makedirs(TASK_DIR, exist_ok=True)


def _task_path(task_id):
    return os.path.join(TASK_DIR, f'{task_id}.json')


def task_set(task_id, **kwargs):
    """更新任务状态，写入文件"""
    p = _task_path(task_id)
    data = {}
    if os.path.exists(p):
        try:
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            data = {}
    data.update(kwargs)
    data['updated_at'] = time.time()
    # output 是文件路径，不能 json 化，单独存
    output = data.pop('_output', None)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    if output:
        with open(p + '.output', 'w', encoding='utf-8') as f:
            f.write(output)


def task_get(task_id):
    """读取任务状态"""
    p = _task_path(task_id)
    if not os.path.exists(p):
        return None
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        out_p = p + '.output'
        if os.path.exists(out_p):
            with open(out_p, 'r', encoding='utf-8') as out_f:
                data['output'] = out_f.read()
        return data
    except Exception:
        return None

# ============== HTML 模板 ==============
INDEX_HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>Lightning Warehouse Calculator - Web</title>
    <style>
        body { font-family: -apple-system, "Microsoft YaHei", sans-serif; max-width: 720px; margin: 40px auto; padding: 0 20px; }
        h1 { color: #2c3e50; }
        .file-row { display: flex; align-items: center; margin: 12px 0; }
        .file-row label { width: 200px; font-weight: bold; }
        .file-row input[type=file] { flex: 1; }
        .file-row .status { width: 220px; color: #888; margin-left: 8px; font-size: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; display: none; }
        .config { background: #f5f5f5; padding: 16px; border-radius: 8px; margin: 20px 0; }
        .config input { width: 100px; padding: 6px; }
        button { background: #3498db; color: white; border: none; padding: 12px 32px; font-size: 16px; border-radius: 6px; cursor: pointer; }
        button:hover { background: #2980b9; }
        button:disabled { background: #95a5a6; cursor: not-allowed; }
        #log { background: #2c3e50; color: #ecf0f1; padding: 16px; border-radius: 6px; font-family: monospace; min-height: 200px; max-height: 400px; overflow-y: auto; white-space: pre-wrap; font-size: 13px; }
        .required { color: #e74c3c; }
        .optional { color: #95a5a6; font-size: 12px; }
    </style>
</head>
<body>
    <h1>Lightning Warehouse Calculator</h1>
    <p>Upload 4 Excel files. Auto pivot + VLOOKUP + smart fill.</p>
    <p><small>文件在服务器上临时处理，处理完成后自动删除。</small></p>

    <form id="uploadForm" enctype="multipart/form-data">
        <div class="file-row">
            <label>1. 闪电仓计算模板 <span class="required">*</span></label>
            <input type="file" name="file1" accept=".xlsx,.xls,.xlsm" required>
            <span class="status" id="s1"></span>
        </div>
        <div class="file-row">
            <label>2. 全店数据-门店成交明细 <span class="required">*</span></label>
            <input type="file" name="file2" accept=".xlsx,.xls,.xlsm,.csv" required>
            <span class="status" id="s2"></span>
        </div>
        <div class="file-row">
            <label>3. 评价分析明细 <span class="optional">(可选)</span></label>
            <input type="file" name="file3" accept=".xlsx,.xls,.xlsm,.csv">
            <span class="status" id="s3"></span>
        </div>
        <div class="file-row">
            <label>4. 门店推广费 <span class="optional">(可选)</span></label>
            <input type="file" name="file4" accept=".xlsx,.xls,.xlsm,.csv">
            <span class="status" id="s4"></span>
        </div>

        <div class="config">
            <label>计算天数: <input type="number" name="days" value="30" min="1" max="365" required></label>
            &nbsp;&nbsp;
            <label>体验分参考: <input type="number" name="experience" value="80" step="1" min="0" max="100"></label>
        </div>

        <button type="submit" id="submitBtn">Start Processing</button>
    </form>

    <h3>Processing Log</h3>
    <div id="log">Waiting...</div>

    <div id="result" style="display:none; margin-top: 20px;">
        <h3>Done</h3>
        <a id="downloadBtn" href="#" download style="background:#27ae60;color:white;padding:12px 24px;text-decoration:none;border-radius:6px;display:inline-block;">Download Result</a>
    </div>

    <script>
        // 显示文件名（纯 ASCII 版本，避免编码问题）
        function bindFileShow() {
            var inputs = document.querySelectorAll('input[type=file]');
            for (var i = 0; i < inputs.length; i++) {
                (function(inp, idx) {
                    inp.addEventListener('change', function(e) {
                        var f = e.target.files[0];
                        var span = document.getElementById('s' + (idx + 1));
                        if (f) {
                            span.textContent = 'OK: ' + f.name;
                            span.style.display = 'inline';
                            span.style.color = '#27ae60';
                        } else {
                            span.textContent = '';
                            span.style.display = 'none';
                        }
                    });
                })(inputs[i], i);
            }
        }

        function startPoll(taskId) {
            var log = document.getElementById('log');
            var result = document.getElementById('result');
            var btn = document.getElementById('submitBtn');
            var timer = setInterval(function() {
                fetch('/status/' + taskId).then(function(r) { return r.json(); }).then(function(d) {
                    log.textContent = d.log || '';
                    log.scrollTop = log.scrollHeight;
                    if (d.status === 'done') {
                        clearInterval(timer);
                        btn.disabled = false;
                        btn.textContent = 'Start Processing';
                        result.style.display = 'block';
                        document.getElementById('downloadBtn').href = '/download/' + taskId;
                    } else if (d.status === 'error') {
                        clearInterval(timer);
                        btn.disabled = false;
                        btn.textContent = 'Start Processing';
                        log.textContent = log.textContent + '\n\nERROR: ' + d.error;
                    }
                }).catch(function(e) {});
            }, 1000);
        }

        function initForm() {
            bindFileShow();
            var form = document.getElementById('uploadForm');
            var btn = document.getElementById('submitBtn');
            var log = document.getElementById('log');
            var result = document.getElementById('result');

            form.addEventListener('submit', function(e) {
                e.preventDefault();
                btn.disabled = true;
                btn.textContent = 'Processing...';
                log.textContent = 'Uploading...';
                result.style.display = 'none';

                var formData = new FormData(form);

                fetch('/process', { method: 'POST', body: formData })
                    .then(function(r) { return r.json(); })
                    .then(function(data) {
                        if (data.task_id) {
                            startPoll(data.task_id);
                        } else {
                            log.textContent = 'Error: ' + (data.error || 'Unknown');
                            btn.disabled = false;
                            btn.textContent = 'Start Processing';
                        }
                    })
                    .catch(function(err) {
                        log.textContent = 'Network Error: ' + err.message;
                        btn.disabled = false;
                        btn.textContent = 'Start Processing';
                    });
            });
        }

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initForm);
        } else {
            initForm();
        }
    </script>
</body>
</html>
"""

# ============== 路由 ==============

@app.route('/')
def index():
    return render_template_string(INDEX_HTML)

@app.route('/health')
def health():
    return jsonify(status='ok', version='1.0')

@app.route('/process', methods=['POST'])
def process():
    """接收文件，启动后台处理"""
    # 检查必填
    for name in ('file1', 'file2', 'days'):
        if name not in request.files and name != 'days':
            return jsonify(error=f'缺少文件: {name}'), 400
    if 'days' not in request.form:
        return jsonify(error='缺少计算天数'), 400

    days = int(request.form['days'])
    experience = int(request.form.get('experience', 80))

    # 保存文件
    task_id = uuid.uuid4().hex[:12]
    work_dir = os.path.join(tempfile.gettempdir(), f'lightning_{task_id}')
    os.makedirs(work_dir, exist_ok=True)

    file1_path = os.path.join(work_dir, f'file1_{request.files["file1"].filename}')
    file2_path = os.path.join(work_dir, f'file2_{request.files["file2"].filename}')
    file1 = request.files['file1']
    file2 = request.files['file2']
    file1.save(file1_path)
    file2.save(file2_path)

    file3_path = file4_path = None
    if 'file3' in request.files and request.files['file3'].filename:
        file3_path = os.path.join(work_dir, f'file3_{request.files["file3"].filename}')
        request.files['file3'].save(file3_path)
    if 'file4' in request.files and request.files['file4'].filename:
        file4_path = os.path.join(work_dir, f'file4_{request.files["file4"].filename}')
        request.files['file4'].save(file4_path)

    # 初始化任务
    task_set(task_id, status='running', log='', work_dir=work_dir)

    # 在后台线程处理
    import threading
    def run():
        try:
            log_lines = []
            def log(msg):
                log_lines.append(msg)
                task_set(task_id, log='\n'.join(log_lines))

            log(f'[1/4] 处理文件2: 全店数据-门店成交明细')
            tmp2 = os.path.join(work_dir, '_pivot_file2.xlsx')
            df2_stores = []
            wb1 = load_workbook(file1_path, data_only=False)
            ws1 = wb1.active
            for row in ws1.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
                v = row[0]
                if v is not None and str(v).strip():
                    df2_stores.append(str(v).strip())
            log(f'  门店数: {len(df2_stores)}')
            make_pivot_file2(file2_path, df2_stores, tmp2, log)
            log(f'  透视完成: {tmp2}')

            tmp3 = None
            if file3_path:
                log(f'[2/4] 处理文件3: 评价分析明细')
                tmp3 = os.path.join(work_dir, '_pivot_file3.xlsx')
                make_pivot_file3(file3_path, df2_stores, tmp3, log)
                log(f'  透视完成: {tmp3}')

            tmp4 = None
            if file4_path:
                log(f'[3/4] 处理文件4: 门店推广费')
                tmp4 = os.path.join(work_dir, '_pivot_file4.xlsx')
                make_pivot_file4(file4_path, df2_stores, tmp4, log)
                log(f'  透视完成: {tmp4}')

            log(f'[4/4] 生成最终结果到文件1')
            output_path = os.path.join(work_dir, '闪电仓计算结果.xlsx')
            process_file1(file1_path, days, tmp2, output_path, log,
                          file3_vlookup_path=tmp3,
                          file4_vlookup_path=tmp4)

            log(f'Done: {output_path}')
            task_set(task_id, status='done', _output=output_path)
        except Exception as e:
            import traceback
            err_msg = str(e) + '\n' + traceback.format_exc()
            log_lines.append(f'\nERROR: {err_msg}')
            task_set(task_id, status='error', error=err_msg, log='\n'.join(log_lines))

    threading.Thread(target=run, daemon=True).start()

    return jsonify(task_id=task_id)

@app.route('/status/<task_id>')
def status(task_id):
    t = task_get(task_id)
    if not t:
        return jsonify(error='任务不存在'), 404
    return jsonify(status=t.get('status', 'running'), log=t.get('log', ''), error=t.get('error'))

@app.route('/download/<task_id>')
def download(task_id):
    t = task_get(task_id)
    if not t or t.get('status') != 'done':
        return jsonify(error='任务未完成'), 400
    out = t.get('output')
    if not out or not os.path.exists(out):
        return jsonify(error='结果文件丢失'), 404
    return send_file(out, as_attachment=True, download_name='result.xlsx')

# ============== 启动 ==============
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
