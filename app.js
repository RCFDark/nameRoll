/**
 * 名字抽选器 - 主逻辑
 * 功能：Excel名单读取、滚动/转盘抽取、动画效果、历史记录
 */

// ===== 全局状态 =====
const state = {
    names: [],                    // 名单列表
    history: [],                  // 抽取历史
    mode: 'scroll',               // 抽取模式: scroll/wheel
    count: 1,                     // 抽取人数
    noRepeat: false,              // 是否去除重复
    isDrawing: false,             // 正在抽取中
    usedNames: []                 // 已抽取过的名字(用于去除重复)
};

// ===== DOM 元素 =====
const elements = {
    fileInput: document.getElementById('fileInput'),
    fileName: document.getElementById('fileName'),
    nameCount: document.getElementById('nameCount'),
    nameList: document.getElementById('nameList'),
    clearNames: document.getElementById('clearNames'),
    modeBtns: document.querySelectorAll('.mode-btn'),
    countBtns: document.querySelectorAll('.count-btn'),
    customCountSection: document.getElementById('customCountSection'),
    customCount: document.getElementById('customCount'),
    noRepeat: document.getElementById('noRepeat'),
    drawBtn: document.getElementById('drawBtn'),
    historyList: document.getElementById('historyList'),
    clearHistory: document.getElementById('clearHistory'),
    resultOverlay: document.getElementById('resultOverlay'),
    resultNames: document.getElementById('resultNames'),
    closeResult: document.getElementById('closeResult'),
    effectsContainer: document.getElementById('effectsContainer')
};

// ===== 常量 =====
const STORAGE_KEY_NAMES = 'nameRoll_names';
const STORAGE_KEY_HISTORY = 'nameRoll_history';
const STORAGE_KEY_USED = 'nameRoll_usedNames';

// ===== 初始化 =====
function init() {
    loadFromStorage();
    bindEvents();
    render();
}

// ===== 事件绑定 =====
function bindEvents() {
    // 文件导入
    elements.fileInput.addEventListener('change', handleFileSelect);
    
    // 清空名单
    elements.clearNames.addEventListener('click', clearNames);
    
    // 模式切换
    elements.modeBtns.forEach(btn => {
        btn.addEventListener('click', () => switchMode(btn.dataset.mode));
    });
    
    // 人数选择
    elements.countBtns.forEach(btn => {
        btn.addEventListener('click', () => selectCount(btn));
    });
    
    // 自定义人数输入
    elements.customCount.addEventListener('change', () => {
        let val = parseInt(elements.customCount.value) || 1;
        val = Math.max(1, Math.min(10, val));
        elements.customCount.value = val;
        state.count = val;
    });
    
    // 去除重复开关
    elements.noRepeat.addEventListener('change', () => {
        state.noRepeat = elements.noRepeat.checked;
        if (!state.noRepeat) {
            state.usedNames = [];
            saveToStorage();
        }
    });
    
    // 开始抽取
    elements.drawBtn.addEventListener('click', startDraw);
    
    // 关闭结果弹窗
    elements.closeResult.addEventListener('click', closeResult);
    elements.resultOverlay.addEventListener('click', (e) => {
        if (e.target === elements.resultOverlay) closeResult();
    });
    
    // 清空历史
    elements.clearHistory.addEventListener('click', clearHistory);
}

// ===== 文件处理 =====
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    elements.fileName.textContent = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 获取第一个sheet
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // 转换为JSON，找出名字列
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            const names = extractNames(jsonData);
            
            if (names.length === 0) {
                alert('未找到名字，请检查Excel文件格式');
                return;
            }
            
            state.names = names;
            state.usedNames = []; // 重置已使用名单
            saveToStorage();
            render();
            
            showMessage(`成功导入 ${names.length} 个名字！`, 'success');
        } catch (error) {
            console.error('读取Excel失败:', error);
            alert('读取Excel文件失败，请确保格式正确');
        }
    };
    reader.readAsArrayBuffer(file);
}

/**
 * 从Excel数据中提取名字
 * 自动识别第一行有数据的列
 */
function extractNames(data) {
    if (!data || data.length === 0) return [];
    
    // 找到非空的列索引
    let nameColIndex = 0;
    const firstRow = data[0];
    
    // 查找第一个非空列
    for (let i = 0; i < firstRow.length; i++) {
        if (firstRow[i] !== undefined && firstRow[i] !== null && String(firstRow[i]).trim() !== '') {
            nameColIndex = i;
            break;
        }
    }
    
    // 提取该列的所有非空名字
    const names = [];
    for (let i = 1; i < data.length; i++) { // 从第2行开始，跳过表头
        const cell = data[i];
        if (cell && cell[nameColIndex] !== undefined && cell[nameColIndex] !== null) {
            const name = String(cell[nameColIndex]).trim();
            if (name) {
                names.push(name);
            }
        }
    }
    
    return names;
}

// ===== 模式切换 =====
function switchMode(mode) {
    state.mode = mode;
    elements.modeBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.mode === mode);
    });
}

// ===== 人数选择 =====
function selectCount(btn) {
    elements.countBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    
    if (btn.dataset.count === 'custom') {
        elements.customCountSection.classList.remove('hidden');
        state.count = parseInt(elements.customCount.value) || 3;
    } else {
        elements.customCountSection.classList.add('hidden');
        state.count = parseInt(btn.dataset.count);
    }
}

// ===== 抽取逻辑 =====
function startDraw() {
    if (state.isDrawing) return;
    
    // 验证名单
    let availableNames = state.names;
    if (state.noRepeat) {
        availableNames = state.names.filter(n => !state.usedNames.includes(n));
    }
    
    if (availableNames.length === 0) {
        alert('名单为空或所有名字都已抽取过！');
        return;
    }
    
    if (state.count > availableNames.length) {
        alert(`抽取人数(${state.count})超过可用人数(${availableNames.length})！`);
        return;
    }
    
    state.isDrawing = true;
    elements.drawBtn.disabled = true;
    elements.drawBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 抽取中...';
    
    // 播放转动音效
    playSpinSound();
    
    // 根据模式执行抽取动画
    if (state.mode === 'scroll') {
        scrollDraw(availableNames);
    } else {
        wheelDraw(availableNames);
    }
}

/**
 * 滚动抽取动画
 */
function scrollDraw(availableNames) {
    const scrollContainer = document.createElement('div');
    scrollContainer.className = 'scroll-container active';
    scrollContainer.innerHTML = '<div class="scroll-names" id="scrollNames">准备抽取...</div>';
    
    // 替换显示区域
    const oldScroll = document.querySelector('.scroll-container');
    if (oldScroll) oldScroll.remove();
    
    elements.drawBtn.parentNode.insertBefore(scrollContainer, elements.drawBtn);
    
    let duration = 3000; // 3秒
    let interval = 100;
    let elapsed = 0;
    
    const scrollInterval = setInterval(() => {
        const randomIndex = Math.floor(Math.random() * availableNames.length);
        document.getElementById('scrollNames').textContent = availableNames[randomIndex];
        elapsed += interval;
        
        if (elapsed >= duration) {
            clearInterval(scrollInterval);
            finalizeDraw(availableNames);
        }
    }, interval);
}

/**
 * 大转盘抽取动画
 */
function wheelDraw(availableNames) {
    // 转盘显示8个"???"
    const wheelNames = Array(8).fill('???');
    
    // 随机抽取结果
    const results = [];
    const tempNames = [...availableNames];
    for (let i = 0; i < state.count; i++) {
        const randomIndex = Math.floor(Math.random() * tempNames.length);
        results.push(tempNames[randomIndex]);
        tempNames.splice(randomIndex, 1);
    }
    
    // 显示转盘
    let wheelContainer = document.querySelector('.wheel-container');
    if (!wheelContainer) {
        wheelContainer = createWheelContainer(wheelNames);
    }
    
    const oldScroll = document.querySelector('.scroll-container');
    if (oldScroll) oldScroll.remove();
    
    wheelContainer.classList.add('active');
    elements.drawBtn.parentNode.insertBefore(wheelContainer, elements.drawBtn);
    
    // 转盘旋转动画
    const wheel = wheelContainer.querySelector('.wheel');
    wheel.style.transition = 'none';
    wheel.style.transform = 'rotate(0deg)';
    wheel.offsetHeight;
    
    // 随机旋转
    const randomRotation = 360 * 5 + Math.random() * 360;
    wheel.style.transition = 'transform 3s cubic-bezier(0.17, 0.67, 0.12, 0.99)';
    wheel.style.transform = `rotate(${randomRotation}deg)`;
    
    // 动画结束后显示结果
    setTimeout(() => {
        state.isDrawing = false;
        elements.drawBtn.disabled = false;
        elements.drawBtn.innerHTML = '<i class="fas fa-hand-pointer"></i> 开始抽取';
        
        wheelContainer.remove();
        showResults(results);
        
        addToHistory(results);
        
        if (state.noRepeat) {
            state.usedNames.push(...results);
            saveToStorage();
        }
    }, 3500);
}

/**
 * 创建转盘容器
 */
function createWheelContainer(wheelNames) {
    const container = document.createElement('div');
    container.className = 'wheel-container';
    container.innerHTML = `
        <div class="wheel-pointer"></div>
        <canvas class="wheel" id="wheelCanvas" width="300" height="300"></canvas>
    `;
    
    setTimeout(() => {
        const canvas = container.querySelector('#wheelCanvas');
        drawWheelCanvas(canvas);
    }, 100);
    
    return container;
}

/**
 * 绘制转盘Canvas
 */
function drawWheelCanvas(canvas) {
    const ctx = canvas.getContext('2d');
    const centerX = canvas.width / 2;
    const centerY = canvas.height / 2;
    const radius = Math.min(centerX, centerY) - 10;
    
    const colors = ['#FF6B6B', '#FFD700', '#87CEEB', '#98FB98', '#FFB6C1', '#DDA0DD', '#FFA500', '#40E0D0'];
    
    const displayNames = Array(8).fill('???');
    const sliceAngle = (2 * Math.PI) / displayNames.length;
    
    // 绘制扇形
    displayNames.forEach((name, i) => {
        ctx.beginPath();
        ctx.moveTo(centerX, centerY);
        ctx.arc(centerX, centerY, radius, i * sliceAngle - Math.PI/2, (i + 1) * sliceAngle - Math.PI/2);
        ctx.closePath();
        ctx.fillStyle = colors[i % colors.length];
        ctx.fill();
        ctx.strokeStyle = '#fff';
        ctx.lineWidth = 2;
        ctx.stroke();
        
        // 绘制文字
        ctx.save();
        ctx.translate(centerX, centerY);
        ctx.rotate(i * sliceAngle + sliceAngle / 2 - Math.PI/2);
        ctx.textAlign = 'right';
        ctx.fillStyle = '#333';
        ctx.font = 'bold 14px "Microsoft YaHei"';
        ctx.fillText(name.substring(0, 4), radius - 20, 5);
        ctx.restore();
    });
    
    // 绘制中心圆
    ctx.beginPath();
    ctx.arc(centerX, centerY, 25, 0, 2 * Math.PI);
    ctx.fillStyle = '#FFD700';
    ctx.fill();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 3;
    ctx.stroke();
    
    // 中心文字
    ctx.fillStyle = '#333';
    ctx.font = 'bold 12px "Microsoft YaHei"';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('抽奖', centerX, centerY);
}

/**
 * 完成抽取，显示结果
 */
function finalizeDraw(availableNames) {
    const results = [];
    const tempNames = [...availableNames];
    
    // 随机抽取指定数量
    for (let i = 0; i < state.count; i++) {
        const randomIndex = Math.floor(Math.random() * tempNames.length);
        results.push(tempNames[randomIndex]);
        tempNames.splice(randomIndex, 1);
    }
    
    // 播放成功音效
    playSuccessSound();
    
    // 显示结果
    showResults(results);
    
    // 记录到历史
    addToHistory(results);
    
    // 更新已使用名单
    if (state.noRepeat) {
        state.usedNames.push(...results);
        saveToStorage();
    }
    
    // 清理滚动容器
    const scrollContainer = document.querySelector('.scroll-container');
    if (scrollContainer) scrollContainer.remove();
    
    state.isDrawing = false;
    elements.drawBtn.disabled = false;
    elements.drawBtn.innerHTML = '<i class="fas fa-play"></i> 开始抽取';
}

/**
 * 显示抽取结果
 */
function showResults(results) {
    // 生成名字HTML
    elements.resultNames.innerHTML = results.map(name => 
        `<span class="name">${escapeHtml(name)}</span>`
    ).join('');
    
    // 显示弹窗
    elements.resultOverlay.classList.remove('hidden');
    
    // 播放特效
    createConfetti();
    createFireworks();
}

function closeResult() {
    elements.resultOverlay.classList.add('hidden');
}

// ===== 历史记录 =====
function addToHistory(results) {
    const record = {
        time: new Date().toLocaleString('zh-CN'),
        names: results,
        mode: state.mode,
        count: state.count
    };
    
    state.history.unshift(record);
    if (state.history.length > 50) { // 最多保存50条
        state.history.pop();
    }
    
    saveToStorage();
    renderHistory();
}

function clearHistory() {
    if (confirm('确定要清空所有抽取历史吗？')) {
        state.history = [];
        saveToStorage();
        renderHistory();
    }
}

// ===== 存储功能 =====
function saveToStorage() {
    try {
        localStorage.setItem(STORAGE_KEY_NAMES, JSON.stringify(state.names));
        localStorage.setItem(STORAGE_KEY_HISTORY, JSON.stringify(state.history));
        localStorage.setItem(STORAGE_KEY_USED, JSON.stringify(state.usedNames));
    } catch (e) {
        console.error('保存到localStorage失败:', e);
    }
}

function loadFromStorage() {
    try {
        const names = localStorage.getItem(STORAGE_KEY_NAMES);
        const history = localStorage.getItem(STORAGE_KEY_HISTORY);
        const usedNames = localStorage.getItem(STORAGE_KEY_USED);
        
        if (names) state.names = JSON.parse(names);
        if (history) state.history = JSON.parse(history);
        if (usedNames) state.usedNames = JSON.parse(usedNames);
    } catch (e) {
        console.error('从localStorage加载失败:', e);
    }
}

function clearNames() {
    if (confirm('确定要清空所有名单吗？')) {
        state.names = [];
        state.usedNames = [];
        elements.fileName.textContent = '未选择文件';
        saveToStorage();
        render();
        showMessage('名单已清空', 'info');
    }
}

// ===== 渲染 =====
function render() {
    // 更新人数统计
    elements.nameCount.textContent = state.names.length;
    
    // 更新名单列表
    renderNameList();
    
    // 更新历史记录
    renderHistory();
}

function renderNameList() {
    if (state.names.length === 0) {
        elements.nameList.innerHTML = '<div class="empty-tip">请导入Excel文件</div>';
        return;
    }
    
    elements.nameList.innerHTML = state.names.map((name, index) => {
        const isUsed = state.usedNames.includes(name);
        return `<div class="name-item" style="${isUsed ? 'opacity: 0.5;' : ''}">
            ${escapeHtml(name)}
            ${isUsed ? '✓' : ''}
        </div>`;
    }).join('');
}

function renderHistory() {
    if (state.history.length === 0) {
        elements.historyList.innerHTML = '<div class="empty-tip">暂无抽取记录</div>';
        return;
    }
    
    elements.historyList.innerHTML = state.history.map(record => `
        <div class="history-item">
            <div class="history-time">${record.time}</div>
            <div class="history-names">
                ${record.names.map(n => `<span>${escapeHtml(n)}</span>`).join('、')}
            </div>
        </div>
    `).join('');
}

// ===== 特效 =====
function createConfetti() {
    const colors = ['#FFD700', '#FFA500', '#FF6B6B', '#87CEEB', '#FFB6C1', '#98FB98'];
    const container = elements.effectsContainer;
    
    for (let i = 0; i < 100; i++) {
        const confetti = document.createElement('div');
        confetti.className = 'confetti';
        confetti.style.left = Math.random() * 100 + '%';
        confetti.style.top = '-10px';
        confetti.style.backgroundColor = colors[Math.floor(Math.random() * colors.length)];
        confetti.style.animationDuration = (Math.random() * 2 + 2) + 's';
        confetti.style.animationDelay = Math.random() * 0.5 + 's';
        
        container.appendChild(confetti);
        
        setTimeout(() => confetti.remove(), 4000);
    }
}

function createFireworks() {
    const colors = ['#FFD700', '#FF6B6B', '#87CEEB', '#98FB98'];
    
    for (let i = 0; i < 5; i++) {
        setTimeout(() => {
            const x = Math.random() * window.innerWidth;
            const y = Math.random() * window.innerHeight * 0.5;
            const color = colors[Math.floor(Math.random() * colors.length)];
            
            for (let j = 0; j < 20; j++) {
                const firework = document.createElement('div');
                firework.className = 'firework';
                firework.style.left = x + 'px';
                firework.style.top = y + 'px';
                firework.style.backgroundColor = color;
                firework.style.animationDuration = (Math.random() * 0.5 + 0.5) + 's';
                
                // 随机方向
                const angle = (Math.PI * 2 * j) / 20;
                const distance = Math.random() * 100 + 50;
                firework.style.setProperty('--tx', Math.cos(angle) * distance + 'px');
                firework.style.setProperty('--ty', Math.sin(angle) * distance + 'px');
                
                elements.effectsContainer.appendChild(firework);
                
                setTimeout(() => firework.remove(), 1000);
            }
        }, i * 200);
    }
}

// ===== 音效 =====
function playSpinSound() {
    // 使用Web Audio API生成简单的音效
    try {
        const audioContext = new (window.AudioContext || window.webkitAudioContext)();
        
        // 创建振荡器模拟转动声音
        const oscillator = audioContext.createOscillator();
        const gainNode = audioContext.createGain();
        
        oscillator.connect(gainNode);
        gainNode.connect(audioContext.destination);
        
        oscillator.type = 'sine';
        oscillator.frequency.setValueAtTime(200, audioContext.currentTime);
        oscillator.frequency.exponentialRampToValueAtTime(400, audioContext.currentTime + 0.1);
        
        gainNode.gain.setValueAtTime(0.1, audioContext.currentTime);
        gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.1);
        
        oscillator.start(audioContext.currentTime);
        oscillator.stop(audioContext.currentTime + 0.1);
    } catch (e) {
        console.log('音效播放失败');
    }
}

function playSuccessSound() {
    try {
        const audioContext = new (window.AudioContext || window.webkitAudioContext)();
        
        // 播放成功的音效序列
        const notes = [523, 659, 784]; // C5, E5, G5
        
        notes.forEach((freq, i) => {
            setTimeout(() => {
                const oscillator = audioContext.createOscillator();
                const gainNode = audioContext.createGain();
                
                oscillator.connect(gainNode);
                gainNode.connect(audioContext.destination);
                
                oscillator.type = 'sine';
                oscillator.frequency.value = freq;
                
                gainNode.gain.setValueAtTime(0.2, audioContext.currentTime);
                gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + 0.3);
                
                oscillator.start(audioContext.currentTime);
                oscillator.stop(audioContext.currentTime + 0.3);
            }, i * 150);
        });
    } catch (e) {
        console.log('音效播放失败');
    }
}

// ===== 工具函数 =====
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function showMessage(text, type = 'info') {
    // 简单的消息提示
    const msg = document.createElement('div');
    msg.style.cssText = `
        position: fixed;
        top: 20px;
        left: 50%;
        transform: translateX(-50%);
        background: ${type === 'success' ? '#4CAF50' : '#2196F3'};
        color: white;
        padding: 12px 24px;
        border-radius: 25px;
        z-index: 2000;
        animation: fadeIn 0.3s ease;
    `;
    msg.textContent = text;
    document.body.appendChild(msg);
    
    setTimeout(() => {
        msg.style.opacity = '0';
        setTimeout(() => msg.remove(), 300);
    }, 2000);
}

// ===== 启动 =====
document.addEventListener('DOMContentLoaded', init);