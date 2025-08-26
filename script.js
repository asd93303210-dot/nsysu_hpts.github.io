// 全局變量
let currentPage = 1;
let testSteps = [];
let stepCounter = 1;

// DOM 元素
const pages = document.querySelectorAll('.page');
const navItems = document.querySelectorAll('.nav-item');
const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');
const submitBtn = document.getElementById('submitBtn');
const addTestStepBtn = document.getElementById('addTestStep');
const testSequence = document.getElementById('testSequence');
const photoInput = document.getElementById('equipmentPhoto');
const photoPreview = document.getElementById('photoPreview');
const timelineCanvas = document.getElementById('timelineCanvas');

// 初始化
document.addEventListener('DOMContentLoaded', function() {
	initializeForm();
	setupEventListeners();
	addDefaultTestStep();
});

// 初始化表單
function initializeForm() {
	showPage(1);
	updateNavigationButtons();
	updateNavigationHighlight();
}

// 設置事件監聽器
function setupEventListeners() {
	prevBtn.addEventListener('click', previousPage);
	nextBtn.addEventListener('click', nextPage);
	addTestStepBtn.addEventListener('click', addTestStep);
	photoInput.addEventListener('change', handlePhotoUpload);
	// 提交→輸出 Word
	submitBtn.addEventListener('click', exportToDocx);
	
	// 頁面導航點擊事件
	navItems.forEach(item => {
		item.addEventListener('click', function() {
			const pageNumber = parseInt(this.getAttribute('data-page'));
			showPage(pageNumber);
		});
	});
}

// 顯示指定頁面
function showPage(pageNumber) {
	pages.forEach((page, index) => {
		if (index + 1 === pageNumber) {
			page.classList.add('active');
		} else {
			page.classList.remove('active');
		}
	});
	currentPage = pageNumber;
	updateNavigationButtons();
	updateNavigationHighlight();
	
	// 如果是第三頁，更新測試步驟
	if (pageNumber === 3) {
		updateTestSteps();
	}
	
	// 如果是第四頁，繪製時序圖
	if (pageNumber === 4) {
		drawTimeline();
	}
}

// 更新導航高亮
function updateNavigationHighlight() {
	navItems.forEach((item, index) => {
		if (index + 1 === currentPage) {
			item.classList.add('active');
		} else {
			item.classList.remove('active');
		}
	});
}

// 更新導航按鈕
function updateNavigationButtons() {
	prevBtn.style.display = currentPage === 1 ? 'none' : 'inline-block';
	nextBtn.style.display = currentPage === 4 ? 'none' : 'inline-block';
	submitBtn.style.display = currentPage === 4 ? 'inline-block' : 'none';
	
	if (currentPage === 1) {
		prevBtn.disabled = true;
	} else {
		prevBtn.disabled = false;
	}
	
	if (currentPage === 4) {
		nextBtn.disabled = true;
	} else {
		nextBtn.disabled = false;
	}
}

// 下一頁
function nextPage() {
	if (currentPage < 4) {
		showPage(currentPage + 1);
	}
}

// 上一頁
function previousPage() {
	if (currentPage > 1) {
		showPage(currentPage - 1);
	}
}

// 添加測試步驟
function addTestStep() {
	const testStep = {
		id: stepCounter++,
		type: 'pressure',
		pressure: '',
		rate: '',
		holdTime: '',
		description: ''
	};
	
	testSteps.push(testStep);
	updateTestSteps();
}

// 移除測試步驟
function removeTestStep(stepId) {
	testSteps = testSteps.filter(step => step.id !== stepId);
	updateTestSteps();
}

// 更新測試步驟顯示
function updateTestSteps() {
	testSequence.innerHTML = '';
	
	testSteps.forEach((step, index) => {
		const stepElement = createTestStepElement(step, index + 1);
		testSequence.appendChild(stepElement);
	});
	
	// 如果沒有測試步驟，添加預設的
	if (testSteps.length === 0) {
		addDefaultTestStep();
	}
}

// 創建測試步驟元素
function createTestStepElement(step, stepNumber) {
	const stepDiv = document.createElement('div');
	stepDiv.className = 'test-step';
	
	// 根據測試類型顯示不同的表單欄位
	let typeSpecificFields = '';
	if (step.type === 'pressure' || step.type === 'depressurize') {
		// 增壓/降壓：顯示速率選擇
		typeSpecificFields = `
			<div class="form-group">
				<label>速率 (bar/分鐘)</label>
				<input type="number" step="0.1" min="0.1" value="${step.rate || ''}" 
				       onchange="updateStepRate(${step.id}, this.value)" placeholder="例如: 1.0">
			</div>
		`;
	} else if (step.type === 'hold') {
		// 持壓：顯示持壓時間選擇
		typeSpecificFields = `
			<div class="form-group">
				<label>持壓時間 (分鐘)</label>
				<input type="number" min="1" value="${step.holdTime || ''}" 
				       onchange="updateStepHoldTime(${step.id}, this.value)" placeholder="例如: 30">
			</div>
		`;
	}
	
	stepDiv.innerHTML = `
		<div class="test-step-header">
			<span class="step-number">${stepNumber}</span>
			<button type="button" class="remove-step" onclick="removeTestStep(${step.id})">移除</button>
		</div>
		<div class="test-step-grid">
			<div class="form-group">
				<label>測試類型</label>
				<select onchange="updateStepType(${step.id}, this.value)">
					<option value="pressure" ${step.type === 'pressure' ? 'selected' : ''}>增壓</option>
					<option value="depressurize" ${step.type === 'depressurize' ? 'selected' : ''}>降壓</option>
					<option value="hold" ${step.type === 'hold' ? 'selected' : ''}>持壓</option>
				</select>
			</div>
			<div class="form-group">
				<label>目標壓力 (bar)</label>
				<input type="number" step="0.1" min="0" value="${step.pressure}" 
				       onchange="updateStepPressure(${step.id}, this.value)" placeholder="例如: 2.5">
			</div>
			${typeSpecificFields}
			<div class="form-group full-width">
				<label>描述</label>
				<textarea rows="2" onchange="updateStepDescription(${step.id}, this.value)" 
				          placeholder="自動生成描述">${generateStepDescription(step)}</textarea>
			</div>
		</div>
	`;
	
	return stepDiv;
}

// 更新步驟類型
function updateStepType(stepId, type) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.type = type;
		// 清除舊的相關欄位
		if (type === 'hold') {
			delete step.rate;
		} else {
			delete step.holdTime;
		}
		updateTestSteps();
	}
}

// 更新步驟速率
function updateStepRate(stepId, rate) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.rate = rate;
		updateTestSteps();
	}
}

// 更新步驟持壓時間
function updateStepHoldTime(stepId, holdTime) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.holdTime = holdTime;
		updateTestSteps();
	}
}

// 更新步驟壓力
function updateStepPressure(stepId, pressure) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.pressure = pressure;
		updateTestSteps();
	}
}

// 更新步驟描述
function updateStepDescription(stepId, description) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.description = description;
	}
}

// 生成步驟描述
function generateStepDescription(step) {
	if (step.type === 'pressure') {
		if (step.pressure && step.rate) {
			return `升壓至${step.pressure} ±0.5 bar，速率${step.rate} bar/分鐘。`;
		}
		return `升壓至${step.pressure || ''} ±0.5 bar，速率${step.rate || ''} bar/分鐘。`;
	} else if (step.type === 'depressurize') {
		if (step.pressure && step.rate) {
			return `降壓至${step.pressure} ±0.5 bar，速率${step.rate} bar/分鐘。`;
		}
		return `降壓至${step.pressure || ''} ±0.5 bar，速率${step.rate || ''} bar/分鐘。`;
	} else if (step.type === 'hold') {
		if (step.pressure && step.holdTime) {
			return `持壓${step.pressure} ±0.5 bar，持續${step.holdTime} 分鐘。`;
		}
		return `持壓${step.pressure || ''} ±0.5 bar，持續${step.holdTime || ''} 分鐘。`;
	}
	return '';
}

// 添加預設測試步驟（只有兩個）
function addDefaultTestStep() {
	const defaultSteps = [
		{ id: stepCounter++, type: 'pressure', pressure: '5.0', rate: '1.0', description: '' },
		{ id: stepCounter++, type: 'depressurize', pressure: '0', rate: '1.0', description: '' }
	];
	
	testSteps = defaultSteps;
	updateTestSteps();
}

// 處理照片上傳
function handlePhotoUpload(event) {
	const file = event.target.files[0];
	if (file) {
		const reader = new FileReader();
		reader.onload = function(e) {
			photoPreview.innerHTML = `<img src="${e.target.result}" alt="設備照片">`;
			photoPreview.style.display = 'block';
		};
		reader.readAsDataURL(file);
	}
}

// 繪製時序圖
function drawTimeline() {
	const ctx = timelineCanvas.getContext('2d');
	const canvas = timelineCanvas;
	
	// 清空畫布
	ctx.clearRect(0, 0, canvas.width, canvas.height);
	
	// 設置樣式
	ctx.font = '12px Microsoft JhengHei';
	ctx.textAlign = 'center';
	
	// 繪製座標軸
	const margin = 60;
	const chartWidth = canvas.width - 2 * margin;
	const chartHeight = canvas.height - 2 * margin;
	
	// 計算總時間和最大壓力
	let totalTime = 0;
	let maxPressure = 0;
	let currentPressure = 0;
	
	// 計算每個步驟的時間和壓力
	const calculatedSteps = [];
	testSteps.forEach(step => {
		if (step.type === 'pressure' || step.type === 'depressurize') {
			if (step.pressure && step.rate) {
				const pressureChange = Math.abs(parseFloat(step.pressure) - currentPressure);
				const stepTime = pressureChange / parseFloat(step.rate);
				calculatedSteps.push({
					...step,
					startTime: totalTime,
					endTime: totalTime + stepTime,
					startPressure: currentPressure,
					endPressure: parseFloat(step.pressure),
					stepTime: stepTime
				});
				totalTime += stepTime;
				currentPressure = parseFloat(step.pressure);
			}
		} else if (step.type === 'hold') {
			if (step.pressure && step.holdTime) {
				const holdTime = parseFloat(step.holdTime);
				calculatedSteps.push({
					...step,
					startTime: totalTime,
					endTime: totalTime + holdTime,
					startPressure: currentPressure,
					endPressure: currentPressure,
					stepTime: holdTime
				});
				totalTime += holdTime;
			}
		}
		
		if (currentPressure > maxPressure) {
			maxPressure = currentPressure;
		}
	});
	
	// 如果沒有有效數據，顯示提示
	if (totalTime === 0 || calculatedSteps.length === 0) {
		ctx.fillStyle = '#666';
		ctx.font = '16px Microsoft JhengHei';
		ctx.fillText('請在第三階段設定測試參數', canvas.width / 2, canvas.height / 2);
		return;
	}
	
	// 調整Y軸範圍
	const pressureRange = Math.max(maxPressure, 12.5);
	
	// Y軸 (壓力)
	ctx.beginPath();
	ctx.moveTo(margin, margin);
	ctx.lineTo(margin, canvas.height - margin);
	ctx.strokeStyle = '#333';
	ctx.lineWidth = 2;
	ctx.stroke();
	
	// X軸 (時間)
	ctx.beginPath();
	ctx.moveTo(margin, canvas.height - margin);
	ctx.lineTo(canvas.width - margin, canvas.height - margin);
	ctx.stroke();
	
	// 繪製壓力標籤
	ctx.fillStyle = '#666';
	for (let i = 0; i <= 5; i++) {
		const y = margin + (chartHeight * (5 - i)) / 5;
		const pressure = (pressureRange * i) / 5;
		ctx.fillText(`${pressure.toFixed(1)} bar`, 30, y + 4);
		
		// 繪製水平參考線
		ctx.beginPath();
		ctx.moveTo(margin, y);
		ctx.lineTo(canvas.width - margin, y);
		ctx.strokeStyle = '#ddd';
		ctx.lineWidth = 1;
		ctx.stroke();
	}
	
	// 繪製時間標籤
	for (let i = 0; i <= 6; i++) {
		const x = margin + (chartWidth * i) / 6;
		const time = Math.round((totalTime * i) / 6);
		ctx.fillText(`${time} min`, x, canvas.height - 30);
		
		// 繪製垂直參考線
		ctx.beginPath();
		ctx.moveTo(x, margin);
		ctx.lineTo(x, canvas.height - margin);
		ctx.strokeStyle = '#ddd';
		ctx.lineWidth = 1;
		ctx.stroke();
	}
	
	// 繪製測試步驟
	ctx.strokeStyle = '#667eea';
	ctx.lineWidth = 3;
	ctx.beginPath();
	
	// 從起始點開始
	ctx.moveTo(margin, canvas.height - margin - (0 / pressureRange) * chartHeight);
	
	calculatedSteps.forEach((step, index) => {
		const x1 = margin + (step.startTime / totalTime) * chartWidth;
		const y1 = canvas.height - margin - (step.startPressure / pressureRange) * chartHeight;
		const x2 = margin + (step.endTime / totalTime) * chartWidth;
		const y2 = canvas.height - margin - (step.endPressure / pressureRange) * chartHeight;
		
		ctx.lineTo(x1, y1);
		ctx.lineTo(x2, y2);
	});
	
	ctx.stroke();
	
	// 繪製數據點和標籤
	calculatedSteps.forEach((step, index) => {
		if (step.type === 'pressure' || step.type === 'depressurize') {
			const x = margin + (step.endTime / totalTime) * chartWidth;
			const y = canvas.height - margin - (step.endPressure / pressureRange) * chartHeight;
			
			// 數據點
			ctx.fillStyle = '#667eea';
			ctx.beginPath();
			ctx.arc(x, y, 4, 0, 2 * Math.PI);
			ctx.fill();
			
			// 壓力標籤
			ctx.fillStyle = '#333';
			ctx.fillText(`${step.endPressure} bar`, x, y - 10);
			
			// 時間標籤
			ctx.fillText(`${step.stepTime.toFixed(1)}min`, x, y + 20);
		} else if (step.type === 'hold') {
			const x = margin + (step.endTime / totalTime) * chartWidth;
			const y = canvas.height - margin - (step.endPressure / pressureRange) * chartHeight;
			
			// 持壓標籤
			ctx.fillStyle = '#333';
			ctx.fillText(`持壓${step.stepTime}min`, x, y + 20);
		}
	});
	
	// 重置樣式
	ctx.strokeStyle = '#333';
	ctx.lineWidth = 2;
}

// 導出為 Word (.docx)
async function exportToDocx() {
	// 確保時序圖為最新
	drawTimeline();
	// 收集資料
	const formData = {
		company: {
			name: document.getElementById('companyName').value || '',
			taxId: document.getElementById('taxId').value || '',
			address: document.getElementById('address').value || '',
			contactPerson: document.getElementById('contactPerson').value || '',
			contactPhone: document.getElementById('contactPhone').value || '',
			contactEmail: document.getElementById('contactEmail').value || ''
		},
		equipment: {
			name: document.getElementById('equipmentName').value || '',
			model: document.getElementById('equipmentModel').value || '',
			specs: document.getElementById('equipmentSpecs').value || '',
			technicalData: document.getElementById('technicalData').value || ''
		},
		testSteps: testSteps
	};
	
	const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ImageRun } = window.docx;
	
	// 標題
	const title = new Paragraph({
		text: '高壓模擬實驗觀測系統使用申請表',
		heading: HeadingLevel.TITLE,
		alignment: AlignmentType.CENTER,
		spacing: { after: 300 }
	});
	
	// 委託單位資料表格
	const companyRows = [
		['單位名稱', formData.company.name],
		['統一編號', formData.company.taxId],
		['地址', formData.company.address],
		['聯絡人', formData.company.contactPerson],
		['聯絡電話', formData.company.contactPhone],
		['E-mail', formData.company.contactEmail]
	].map(([k, v]) => new TableRow({
		children: [
			new TableCell({ children: [new Paragraph(k)] }),
			new TableCell({ children: [new Paragraph(v)] })
		]
	}));
	const companyTable = new Table({
		rows: companyRows,
		width: { size: 100, type: WidthType.PERCENTAGE }
	});
	
	// 設備資料表格
	const equipRows = [
		['設備名稱', formData.equipment.name],
		['型號', formData.equipment.model],
		['尺寸及規格', formData.equipment.specs],
		['其它技術資料', formData.equipment.technicalData]
	].map(([k, v]) => new TableRow({
		children: [
			new TableCell({ children: [new Paragraph(k)] }),
			new TableCell({ children: [new Paragraph(v)] })
		]
	}));
	const equipTable = new Table({
		rows: equipRows,
		width: { size: 100, type: WidthType.PERCENTAGE }
	});
	
	// 測試要求（編號清單）
	const stepParas = [];
	formData.testSteps.forEach((s, idx) => {
		const desc = s.description && s.description.trim().length > 0 ? s.description : generateStepDescription(s);
		stepParas.push(new Paragraph({ text: `${idx + 1}. ${desc}`, spacing: { after: 120 } }));
	});
	
	// 先嘗試讀取Canvas圖片為byte陣列
	let timelineImageRun = null;
	try {
		const dataUrl = timelineCanvas.toDataURL('image/png');
		const imageData = dataURLToUint8Array(dataUrl);
		timelineImageRun = new ImageRun({ data: imageData, transformation: { width: 600, height: 300 } });
	} catch (e) {
		// ignore
	}
	
	// 章節標題
	const h1 = (t) => new Paragraph({ text: t, heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 200 } });
	
	// 建立文件
	const docChildren = [
		title,
		h1('一、委託單位資料'),
		companyTable,
		h1('二、擬測試設備資料'),
		equipTable,
		h1('三、壓力測試要求'),
		...stepParas,
		h1('四、測試時序圖')
	];
	// 若有圖片，立即在children中加入圖片段落
	if (timelineImageRun) {
		docChildren.push(new Paragraph({ children: [timelineImageRun], alignment: AlignmentType.CENTER }));
	}

	const doc = new Document({
		sections: [{ children: docChildren }]
	});
	
	const blob = await Packer.toBlob(doc);
	const fileName = `高壓艙申請表_${new Date().toISOString().slice(0,10)}.docx`;
	triggerDownload(blob, fileName);
}

function dataURLToUint8Array(dataURL) {
	const base64 = dataURL.split(',')[1];
	const binary = atob(base64);
	const len = binary.length;
	const bytes = new Uint8Array(len);
	for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
	return bytes;
}

function triggerDownload(blob, fileName) {
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = fileName;
	document.body.appendChild(a);
	a.click();
	setTimeout(() => {
		document.body.removeChild(a);
		URL.revokeObjectURL(url);
	}, 0);
}
