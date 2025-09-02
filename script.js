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
	// 生成時序圖
	document.getElementById('generateTimelineBtn').addEventListener('click', generateTimeline);
	
	// 頁面導航點擊事件
	navItems.forEach(item => {
		item.addEventListener('click', function() {
			const pageNumber = parseInt(this.getAttribute('data-page'));
			showPage(pageNumber);
		});
	});
	
	// 接線要求勾選框事件
	setupWiringCheckboxes();
	
	// 系統選擇事件
	setupSystemSelection();
	
	// 價格試算事件
	setupPricingCalculation();
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
	
	// 如果是第四頁，更新測試步驟
	if (pageNumber === 4) {
		updateTestSteps();
	}
	
	// 如果是第六頁，計算價格
	if (pageNumber === 6) {
		calculatePricing();
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
	nextBtn.style.display = currentPage === 6 ? 'none' : 'inline-block';
	submitBtn.style.display = currentPage === 6 ? 'inline-block' : 'none';
	
	if (currentPage === 1) {
		prevBtn.disabled = true;
	} else {
		prevBtn.disabled = false;
	}
	
	if (currentPage === 6) {
		nextBtn.disabled = true;
	} else {
		nextBtn.disabled = false;
	}
}

// 下一頁
function nextPage() {
	if (currentPage < 6) {
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
				<label>速率 (bar/min)</label>
				<input type="number" step="0.1" min="0.1" value="${step.rate || ''}" 
				       onchange="updateStepRate(${step.id}, this.value)" placeholder="例如: 1.0">
			</div>
		`;
	} else if (step.type === 'hold') {
		// 持壓：顯示持壓時間選擇
		typeSpecificFields = `
			<div class="form-group">
				<label>持壓時間 (min)</label>
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
			return `增壓至${step.pressure} ±0.5 bar，速率${step.rate} bar/min。`;
		}
		return `增壓至${step.pressure || ''} ±0.5 bar，速率${step.rate || ''} bar/min。`;
	} else if (step.type === 'depressurize') {
		if (step.pressure && step.rate) {
			return `降壓至${step.pressure} ±0.5 bar，速率${step.rate} bar/min。`;
		}
		return `降壓至${step.pressure || ''} ±0.5 bar，速率${step.rate || ''} bar/min。`;
	} else if (step.type === 'hold') {
		if (step.pressure && step.holdTime) {
			return `持壓${step.pressure} ±0.5 bar，持續${step.holdTime} min。`;
		}
		return `持壓${step.pressure || ''} ±0.5 bar，持續${step.holdTime || ''} min。`;
	}
	return '';
}

// 添加預設測試步驟（三個：增壓、持壓、降壓）
function addDefaultTestStep() {
	const defaultSteps = [
		{ id: stepCounter++, type: 'pressure', pressure: '5.0', rate: '1.0', description: '' },
		{ id: stepCounter++, type: 'hold', pressure: '5.0', holdTime: '30', description: '' },
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

// 生成時序圖
function generateTimeline() {
	// 檢查是否有有效的測試步驟
	const validSteps = testSteps.filter(step => {
		if (step.type === 'pressure' || step.type === 'depressurize') {
			return step.pressure && step.rate;
		} else if (step.type === 'hold') {
			return step.pressure && step.holdTime;
		}
		return false;
	});
	
	if (validSteps.length === 0) {
		alert('請先填寫完整的測試步驟參數');
		return;
	}
	
	// 繪製時序圖
	drawTimeline();
	
	// 顯示時序圖容器
	document.getElementById('timelineContainer').style.display = 'block';
	
	// 滾動到時序圖區域
	document.getElementById('timelineContainer').scrollIntoView({ behavior: 'smooth' });
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
		ctx.fillText('請在第四階段設定測試參數', canvas.width / 2, canvas.height / 2);
		return;
	}
	
	// 調整Y軸範圍
	const pressureRange = Math.max(maxPressure, 12.5);
	
	// 繪製壓力標籤（Y軸刻度）
	ctx.fillStyle = '#000';
	ctx.textAlign = 'right';
	for (let i = 0; i <= 5; i++) {
		const y = margin + (chartHeight * (5 - i)) / 5;
		const pressure = (pressureRange * i) / 5;
		ctx.fillText(`${pressure.toFixed(1)}`, margin - 10, y + 4);
	}
	
	// 繪製時間標籤（X軸刻度）
	ctx.textAlign = 'center';
	for (let i = 0; i <= 6; i++) {
		const x = margin + (chartWidth * i) / 6;
		const time = Math.round((totalTime * i) / 6);
		ctx.fillText(`${time}`, x, canvas.height - margin + 20);
	}
	
	// 繪製軸標題
	ctx.font = '14px Microsoft JhengHei';
	ctx.fillStyle = '#000';
	
	// Y軸標題：壓力 (bar)
	ctx.save();
	ctx.translate(20, canvas.height / 2);
	ctx.rotate(-Math.PI / 2);
	ctx.fillText('壓力 (bar)', 0, 0);
	ctx.restore();
	
	// X軸標題：時間 (min)
	ctx.fillText('時間 (min)', canvas.width / 2, canvas.height - 10);
	
	// 重置字體大小
	ctx.font = '12px Microsoft JhengHei';
	
	// 繪製格線（淺灰色）
	ctx.strokeStyle = '#e0e0e0';
	ctx.lineWidth = 1;
	
	// 水平格線（壓力線）
	for (let i = 1; i < 5; i++) { // 跳過最外側的兩條線
		const y = margin + (chartHeight * (5 - i)) / 5;
		ctx.beginPath();
		ctx.moveTo(margin, y);
		ctx.lineTo(canvas.width - margin, y);
		ctx.stroke();
	}
	
	// 垂直格線（時間線）
	for (let i = 1; i < 6; i++) { // 跳過最外側的兩條線
		const x = margin + (chartWidth * i) / 6;
		ctx.beginPath();
		ctx.moveTo(x, margin);
		ctx.lineTo(x, canvas.height - margin);
		ctx.stroke();
	}
	
	// 繪製座標軸（黑色粗線）
	ctx.strokeStyle = '#000';
	ctx.lineWidth = 2;
	
	// Y軸 (壓力)
	ctx.beginPath();
	ctx.moveTo(margin, margin);
	ctx.lineTo(margin, canvas.height - margin);
	ctx.stroke();
	
	// X軸 (時間)
	ctx.beginPath();
	ctx.moveTo(margin, canvas.height - margin);
	ctx.lineTo(canvas.width - margin, canvas.height - margin);
	ctx.stroke();
	
	// 繪製最外側的格線（與座標軸一致的黑色粗線）
	// 最上方的水平線
	ctx.beginPath();
	ctx.moveTo(margin, margin);
	ctx.lineTo(canvas.width - margin, margin);
	ctx.stroke();
	
	// 最右側的垂直線
	ctx.beginPath();
	ctx.moveTo(canvas.width - margin, margin);
	ctx.lineTo(canvas.width - margin, canvas.height - margin);
	ctx.stroke();
	
	// 繪製測試步驟曲線
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
	
	// 重置樣式
	ctx.strokeStyle = '#000';
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
		testSteps: testSteps,
		wiring: collectWiringData()
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
	
	// 接線要求內容
	const wiringParas = [];
	if (formData.wiring.selectedSystem) {
		// 添加選中系統的規格說明
		wiringParas.push(new Paragraph({ 
			text: '選用系統規格說明：', 
			spacing: { before: 200, after: 100 },
			heading: HeadingLevel.HEADING_2
		}));
		
		if (formData.wiring.selectedSystem === '450') {
			wiringParas.push(new Paragraph({ 
				text: 'HPT-450-230L - 最大操作壓力：450 bar，艙體空間尺寸：Ø 460 mm × L 600 mm', 
				spacing: { after: 120 } 
			}));
			
			if (formData.wiring.system450.length > 0) {
				wiringParas.push(new Paragraph({ text: 'HPT-450-230L 系統接線要求：', spacing: { before: 200, after: 100 } }));
				formData.wiring.system450.forEach(item => {
					const componentText = item.componentSpec ? ` - 接頭零件規格：${item.componentSpec}` : '';
					const remarkText = item.remark ? ` - 備註：${item.remark}` : '';
					wiringParas.push(new Paragraph({ 
						text: `(${item.port}) ${item.spec}${item.description ? ': ' + item.description : ''}${componentText}${remarkText}`, 
						spacing: { after: 80 } 
					}));
				});
			}
		} else if (formData.wiring.selectedSystem === '800') {
			wiringParas.push(new Paragraph({ 
				text: 'HPT-800-85L - 最大操作壓力：800 bar，艙體空間尺寸：Ø 250 mm × L 800 mm', 
				spacing: { after: 120 } 
			}));
			
			if (formData.wiring.system800.length > 0) {
				wiringParas.push(new Paragraph({ text: 'HPT-800-85L 系統接線要求：', spacing: { before: 200, after: 100 } }));
				formData.wiring.system800.forEach(item => {
					const componentText = item.componentSpec ? ` - 接頭零件規格：${item.componentSpec}` : '';
					const remarkText = item.remark ? ` - 備註：${item.remark}` : '';
					wiringParas.push(new Paragraph({ 
						text: `(${item.port}) ${item.spec}${item.description ? ': ' + item.description : ''}${componentText}${remarkText}`, 
						spacing: { after: 80 } 
					}));
				});
			}
		}
	}
	
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
	
	// 添加接線要求
	if (wiringParas.length > 0) {
		docChildren.push(h1('五、接線要求'));
		docChildren.push(...wiringParas);
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

// 設置接線要求勾選框事件
function setupWiringCheckboxes() {
	const checkboxes = document.querySelectorAll('.port-option input[type="checkbox"]:not([disabled])');
	
	checkboxes.forEach(checkbox => {
		checkbox.addEventListener('change', function() {
			const portOption = this.closest('.port-option');
			const inputGroup = portOption.querySelector('.input-group');
			
			if (this.checked) {
				inputGroup.style.display = 'flex';
				// 聚焦到第一個輸入框
				const firstInput = inputGroup.querySelector('.component-input');
				if (firstInput) {
					firstInput.focus();
				}
			} else {
				inputGroup.style.display = 'none';
				// 清空所有輸入框
				const inputs = inputGroup.querySelectorAll('.component-input');
				inputs.forEach(input => {
					input.value = '';
				});
			}
		});
	});
}

// 設置系統選擇事件
function setupSystemSelection() {
	const systemRadios = document.querySelectorAll('input[name="selectedSystem"]');
	
	systemRadios.forEach(radio => {
		radio.addEventListener('change', function() {
			const selectedSystem = this.value;
			const system450Section = document.getElementById('system450Section');
			const system800Section = document.getElementById('system800Section');
			
			// 隱藏所有系統選項
			system450Section.style.display = 'none';
			system800Section.style.display = 'none';
			
			// 清除所有選項的勾選狀態（除了鎖定的(1)(2)）
			clearAllPortSelections();
			
			// 顯示選中的系統
			if (selectedSystem === '450') {
				system450Section.style.display = 'block';
			} else if (selectedSystem === '800') {
				system800Section.style.display = 'block';
			}
		});
	});
}

// 清除所有端口選擇（除了鎖定的(1)(2)）
function clearAllPortSelections() {
	const allCheckboxes = document.querySelectorAll('.port-option input[type="checkbox"]:not([disabled])');
	allCheckboxes.forEach(checkbox => {
		checkbox.checked = false;
		const portOption = checkbox.closest('.port-option');
		const inputGroup = portOption.querySelector('.input-group');
		if (inputGroup) {
			inputGroup.style.display = 'none';
			const inputs = inputGroup.querySelectorAll('.component-input');
			inputs.forEach(input => {
				input.value = '';
			});
		}
	});
}

// 收集接線要求資料
function collectWiringData() {
	const wiringData = {
		selectedSystem: null,
		system450: [],
		system800: []
	};
	
	// 獲取選中的系統
	const selectedSystemRadio = document.querySelector('input[name="selectedSystem"]:checked');
	if (selectedSystemRadio) {
		wiringData.selectedSystem = selectedSystemRadio.value;
	}
	
	// 只收集選中系統的資料
	if (wiringData.selectedSystem === '450') {
		const system450Checkboxes = document.querySelectorAll('input[data-system="450"]');
		system450Checkboxes.forEach(checkbox => {
			if (checkbox.checked) {
				const portOption = checkbox.closest('.port-option');
				const inputGroup = portOption.querySelector('.input-group');
				const inputs = inputGroup.querySelectorAll('.component-input');
				
				wiringData.system450.push({
					port: checkbox.getAttribute('data-port'),
					spec: checkbox.getAttribute('data-spec'),
					description: checkbox.getAttribute('data-description'),
					componentSpec: inputs[0] ? inputs[0].value : '',
					remark: inputs[1] ? inputs[1].value : ''
				});
			}
		});
	} else if (wiringData.selectedSystem === '800') {
		const system800Checkboxes = document.querySelectorAll('input[data-system="800"]');
		system800Checkboxes.forEach(checkbox => {
			if (checkbox.checked) {
				const portOption = checkbox.closest('.port-option');
				const inputGroup = portOption.querySelector('.input-group');
				const inputs = inputGroup.querySelectorAll('.component-input');
				
				wiringData.system800.push({
					port: checkbox.getAttribute('data-port'),
					spec: checkbox.getAttribute('data-spec'),
					description: checkbox.getAttribute('data-description'),
					componentSpec: inputs[0] ? inputs[0].value : '',
					remark: inputs[1] ? inputs[1].value : ''
				});
			}
		});
	}
	
	return wiringData;
}

// 設置價格試算功能
function setupPricingCalculation() {
	// 額外服務勾選框事件
	const englishReportCheckbox = document.getElementById('englishReport');
	const videoRecordingCheckbox = document.getElementById('videoRecording');
	
	if (englishReportCheckbox) {
		englishReportCheckbox.addEventListener('change', calculatePricing);
	}
	
	if (videoRecordingCheckbox) {
		videoRecordingCheckbox.addEventListener('change', calculatePricing);
	}
}

// 計算價格
function calculatePricing() {
	// 獲取測試參數
	const testParams = getTestParameters();
	
	if (!testParams.maxPressure || !testParams.totalTime) {
		// 如果沒有測試參數，顯示預設值
		updateTestParamsDisplay('-- bar', '-- 小時', '--');
		updatePricingDisplay(0, 0, 0, 0, 0, 0, 0);
		return;
	}
	
	// 更新測試參數顯示
	const totalMinutes = Math.round(testParams.totalTime);
	const hours = Math.floor(totalMinutes / 60);
	const minutes = totalMinutes % 60;
	const timeDisplay = minutes > 0 ? `${hours}小時${minutes}分鐘` : `${hours}小時`;
	updateTestParamsDisplay(
		`${testParams.maxPressure} bar`,
		timeDisplay,
		testParams.systemName
	);
	
	// 計算基本費用（直接使用分鐘）
	const pricing = calculateBasePricing(testParams.maxPressure, testParams.totalTime);
	
	// 計算額外服務費用
	const additionalFees = calculateAdditionalFees(pricing.basicTestFee);
	
	// 更新價格顯示
	updatePricingDisplay(
		pricing.baseFee,
		pricing.overtimeFee,
		pricing.overtimeMinutes,
		pricing.extendedOvertimeFee,
		pricing.extendedOvertimeMinutes,
		additionalFees.englishReport,
		additionalFees.videoRecording
	);
	
	// 計算總計
	const total = pricing.basicTestFee + additionalFees.englishReport + additionalFees.videoRecording;
	updateTotalDisplay(pricing.basicTestFee, additionalFees.englishReport, additionalFees.videoRecording, total);
}

// 獲取測試參數
function getTestParameters() {
	// 從第四階段獲取最大壓力和總時間
	let maxPressure = 0;
	let totalTime = 0;
	let currentPressure = 0; // 追蹤當前壓力
	
	// 計算測試步驟的總時間和最大壓力
	testSteps.forEach(step => {
		if (step.type === 'pressure' || step.type === 'depressurize') {
			if (step.pressure && step.rate) {
				const pressureChange = Math.abs(parseFloat(step.pressure) - currentPressure);
				const stepTime = pressureChange / parseFloat(step.rate);
				totalTime += stepTime;
				maxPressure = Math.max(maxPressure, parseFloat(step.pressure));
				currentPressure = parseFloat(step.pressure); // 更新當前壓力
			}
		} else if (step.type === 'hold') {
			if (step.holdTime) {
				totalTime += parseFloat(step.holdTime); // 持壓時間已經是分鐘
				maxPressure = Math.max(maxPressure, parseFloat(step.pressure));
			}
		}
	});
	
	// 獲取選中的系統
	const selectedSystemRadio = document.querySelector('input[name="selectedSystem"]:checked');
	const systemName = selectedSystemRadio ? (selectedSystemRadio.value === '450' ? 'HPT-450-230L' : 'HPT-800-85L') : '--';
	
	return {
		maxPressure: maxPressure,
		totalTime: totalTime,
		systemName: systemName
	};
}

// 計算基本定價
function calculateBasePricing(maxPressure, totalTimeMinutes) {
	// 根據壓力範圍確定基本費用和超時費用
	let baseFee, overtimeRatePerHour;
	
	if (maxPressure <= 100) {
		baseFee = 45000;
		overtimeRatePerHour = 5000;
	} else if (maxPressure <= 200) {
		baseFee = 55000;
		overtimeRatePerHour = 6000;
	} else if (maxPressure <= 450) {
		baseFee = 70000;
		overtimeRatePerHour = 8000;
	} else if (maxPressure <= 800) {
		baseFee = 90000;
		overtimeRatePerHour = 10000;
	} else {
		baseFee = 90000; // 超過800 bar按最高級別計算
		overtimeRatePerHour = 10000;
	}
	
	// 計算超時費用（使用分鐘為單位）
	let overtimeFee = 0;
	let overtimeMinutes = 0;
	let extendedOvertimeFee = 0;
	let extendedOvertimeMinutes = 0;
	
	const threeHoursInMinutes = 3 * 60; // 180分鐘
	const eightHoursInMinutes = 8 * 60; // 480分鐘
	
	if (totalTimeMinutes > threeHoursInMinutes) {
		// 第4-8小時的超時費用（最多5小時 = 300分鐘）
		const regularOvertimeMinutes = Math.min(totalTimeMinutes - threeHoursInMinutes, 5 * 60);
		overtimeMinutes = regularOvertimeMinutes;
		overtimeFee = Math.round((regularOvertimeMinutes / 60) * overtimeRatePerHour);
		
		// 第9小時起的延長超時費用（1.5倍）
		if (totalTimeMinutes > eightHoursInMinutes) {
			extendedOvertimeMinutes = totalTimeMinutes - eightHoursInMinutes;
			extendedOvertimeFee = Math.round((extendedOvertimeMinutes / 60) * overtimeRatePerHour * 1.5);
		}
	}
	
	const basicTestFee = baseFee + overtimeFee + extendedOvertimeFee;
	
	return {
		baseFee: baseFee,
		overtimeFee: overtimeFee,
		overtimeMinutes: overtimeMinutes,
		extendedOvertimeFee: extendedOvertimeFee,
		extendedOvertimeMinutes: extendedOvertimeMinutes,
		basicTestFee: basicTestFee
	};
}

// 計算額外服務費用
function calculateAdditionalFees(basicTestFee) {
	const englishReport = document.getElementById('englishReport').checked ? 12000 : 0;
	const videoRecording = document.getElementById('videoRecording').checked ? Math.round(basicTestFee * 0.05) : 0;
	
	return {
		englishReport: englishReport,
		videoRecording: videoRecording
	};
}

// 更新測試參數顯示
function updateTestParamsDisplay(maxPressure, totalTime, systemName) {
	document.getElementById('maxPressure').textContent = maxPressure;
	document.getElementById('totalTime').textContent = totalTime;
	document.getElementById('selectedSystemName').textContent = systemName;
}

// 更新價格計算顯示
function updatePricingDisplay(baseFee, overtimeFee, overtimeMinutes, extendedOvertimeFee, extendedOvertimeMinutes, englishReport, videoRecording) {
	// 基本費用
	document.getElementById('baseFee').textContent = `NT$ ${baseFee.toLocaleString()}`;
	
	// 超時費用
	const overtimeSection = document.getElementById('overtimeSection');
	const overtimeFeeElement = document.getElementById('overtimeFee');
	const overtimeHoursElement = document.getElementById('overtimeHours');
	
	if (overtimeMinutes > 0) {
		overtimeSection.style.display = 'block';
		overtimeFeeElement.textContent = `NT$ ${overtimeFee.toLocaleString()}`;
		const overtimeH = Math.floor(overtimeMinutes / 60);
		const overtimeM = overtimeMinutes % 60;
		const overtimeDisplay = overtimeM > 0 ? `${overtimeH}小時${overtimeM}分鐘` : `${overtimeH}小時`;
		overtimeHoursElement.textContent = overtimeDisplay;
	} else {
		overtimeSection.style.display = 'none';
	}
	
	// 延長超時費用
	const extendedOvertimeSection = document.getElementById('extendedOvertimeSection');
	const extendedOvertimeFeeElement = document.getElementById('extendedOvertimeFee');
	const extendedOvertimeHoursElement = document.getElementById('extendedOvertimeHours');
	
	if (extendedOvertimeMinutes > 0) {
		extendedOvertimeSection.style.display = 'block';
		extendedOvertimeFeeElement.textContent = `NT$ ${extendedOvertimeFee.toLocaleString()}`;
		const extendedH = Math.floor(extendedOvertimeMinutes / 60);
		const extendedM = extendedOvertimeMinutes % 60;
		const extendedDisplay = extendedM > 0 ? `${extendedH}小時${extendedM}分鐘` : `${extendedH}小時`;
		extendedOvertimeHoursElement.textContent = extendedDisplay;
	} else {
		extendedOvertimeSection.style.display = 'none';
	}
}

// 更新總計顯示
function updateTotalDisplay(basicTestFee, englishReport, videoRecording, total) {
	// 基本測試費用
	document.getElementById('basicTestFee').textContent = `NT$ ${basicTestFee.toLocaleString()}`;
	
	// 英文報告費用
	const englishReportFee = document.getElementById('englishReportFee');
	if (englishReport > 0) {
		englishReportFee.style.display = 'flex';
	} else {
		englishReportFee.style.display = 'none';
	}
	
	// 錄影服務費用
	const videoRecordingFee = document.getElementById('videoRecordingFee');
	const videoFeeAmount = document.getElementById('videoFeeAmount');
	if (videoRecording > 0) {
		videoRecordingFee.style.display = 'flex';
		videoFeeAmount.textContent = `NT$ ${videoRecording.toLocaleString()}`;
	} else {
		videoRecordingFee.style.display = 'none';
	}
	
	// 總計
	document.getElementById('finalTotal').textContent = `NT$ ${total.toLocaleString()}`;
}
