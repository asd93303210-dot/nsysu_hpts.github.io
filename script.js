// 全局變量
let currentPage = 1;
let testSteps = [];
let stepCounter = 1;
let testObjects = [];
let testObjectCounter = 1;

// DOM 元素變數（將在DOMContentLoaded時初始化）
let pages, navItems, prevBtn, nextBtn, submitBtn, addTestStepBtn, testSequence, timelineCanvas, addTestObjectBtn, testObjectsList;
let ackNotice;
let ackInlineContainer;
let ackLabelText;

// 初始化
document.addEventListener('DOMContentLoaded', function() {
	// 初始化DOM元素引用
	pages = document.querySelectorAll('.page');
	navItems = document.querySelectorAll('.nav-item');
	prevBtn = document.getElementById('prevBtn');
	nextBtn = document.getElementById('nextBtn');
	submitBtn = document.getElementById('submitBtn');
	addTestStepBtn = document.getElementById('addTestStep');
	testSequence = document.getElementById('testSequence');
	timelineCanvas = document.getElementById('timelineCanvas');
	addTestObjectBtn = document.getElementById('addTestObject');
	testObjectsList = document.getElementById('testObjectsList');
	ackNotice = document.getElementById('ackNotice');
	ackInlineContainer = document.getElementById('ackInlineContainer');
	ackLabelText = document.getElementById('ackLabelText');
	
	
	
	initializeForm();
	setupEventListeners();
	addDefaultTestStep();
	addDefaultTestObject();
});

// 將數值格式化為一位小數（字串），空值則回空字串
function formatToOneDecimal(value) {
	if (value === null || value === undefined || value === '') return '';
	const num = parseFloat(value);
	if (isNaN(num)) return '';
	return (Math.round(num * 10) / 10).toFixed(1);
}

// 初始化表單
function initializeForm() {
	showPage(1);
	updateNavigationButtons();
	updateNavigationHighlight();
	
	// 初始化接線要求
	initializeWiring();
}

// 設置事件監聽器
function setupEventListeners() {
	prevBtn.addEventListener('click', previousPage);
	nextBtn.addEventListener('click', nextPage);
	addTestStepBtn.addEventListener('click', addTestStep);
	addTestObjectBtn.addEventListener('click', addTestObject);
	
	// 提交→輸出 Word
	submitBtn.addEventListener('click', exportToDocx);
	
	// 頁面導航點擊事件
	navItems.forEach((item, index) => {
		item.addEventListener('click', function() {
			const pageNumber = parseInt(this.getAttribute('data-page'));
			// 未勾選時禁止從第1頁跳轉到第2頁以上
			if (pageNumber >= 2 && currentPage === 1 && !(ackNotice && ackNotice.checked)) {
				alert('請先詳讀「注意事項」，並於頁尾勾選同意後再進行下一步。');
				return;
			}
			// 第2、3階段必填檢查：從當前頁嘗試跳至更後頁面時
			if (currentPage === 2 && pageNumber > 2) {
				if (!validatePage2Required()) {
					alert('請先完成第2階段之必填欄位（標示*），再進行下一步。');
					return;
				}
			}
			if (currentPage === 3 && pageNumber > 3) {
				if (!validatePage3Required()) {
					alert('請先完成第3階段之必填欄位（標示*），再進行下一步。');
					return;
				}
			}
			showPage(pageNumber);
		});
	});
	
	// 接線要求勾選框事件
	setupWiringCheckboxes();
	
	// 系統選擇事件
	setupSystemSelection();
	
	// 價格試算事件
	setupPricingCalculation();
	
	// 第一頁閱讀確認事件
	if (ackNotice) {
		ackNotice.addEventListener('change', function() {
			updateNavigationButtons();
		});
	}
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
	
	// 如果是第三頁，更新受測物顯示
	if (pageNumber === 3) {
		updateTestObjects();
	}
	
	// 如果是第四頁，更新測試步驟
	if (pageNumber === 4) {
		updateTestSteps();
	}
	
	// 如果是第五頁，套用系統選擇限制
	if (pageNumber === 5) {
		enforceSystemEligibility();
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
		// 讓下一步可點擊以便顯示提醒
		nextBtn.disabled = false;
	} else if (currentPage === 6) {
		prevBtn.disabled = false;
		nextBtn.disabled = true;
	} else {
		prevBtn.disabled = false;
		nextBtn.disabled = false;
	}
}

// 下一頁
function nextPage() {
	// 第一頁需已勾選閱讀確認
	if (currentPage === 1 && !(ackNotice && ackNotice.checked)) {
		alert('請先詳讀「注意事項」，並於頁尾勾選同意後再進行下一步。');
		return;
	}
	// 第二頁必填檢查
	if (currentPage === 2) {
		if (!validatePage2Required()) {
			alert('請先完成第2階段之必填欄位（標示*），再進行下一步。');
			return;
		}
	}
	// 第三頁必填檢查
	if (currentPage === 3) {
		if (!validatePage3Required()) {
			alert('請先完成第3階段之必填欄位（標示*），再進行下一步。');
			return;
		}
	}
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
	const prev = testSteps[testSteps.length - 1] || {};
	const testStep = {
		id: stepCounter++,
		type: prev.type || 'pressure',
		pressure: prev.pressure || '',
		rate: prev.rate || '',
		holdTime: prev.holdTime || '',
		description: ''
	};
	
	testSteps.push(testStep);
	updateTestSteps();
	// 立即更新一次時序圖
	autoUpdateTimeline();
}

// 移除測試步驟
function removeTestStep(stepId) {
	testSteps = testSteps.filter(step => step.id !== stepId);
	updateTestSteps();
	autoUpdateTimeline();
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
	
	// 任何重繪後都觸發時序圖自動更新
	autoUpdateTimeline();
	// 根據最大壓力限制第5階段系統選擇
	enforceSystemEligibility();
}

// 創建測試步驟元素
function createTestStepElement(step, stepNumber) {
	const stepDiv = document.createElement('div');
	stepDiv.className = 'test-step';

	// 取得前一步的目標壓力（用於持壓同步）
	const idx = testSteps.indexOf(step);
	const prevStep = idx > 0 ? testSteps[idx - 1] : null;
	
	// 根據測試類型顯示不同的表單欄位
	let typeSpecificFields = '';
	if (step.type === 'pressure' || step.type === 'depressurize') {
		// 增壓/降壓：顯示速率選擇（下拉）
		const buildRateOptions = (min, max, stepVal, current) => {
			const opts = [];
			for (let v = min; v <= max + 1e-9; v = Math.round((v + stepVal) * 100) / 100) {
				const val = v.toFixed(2);
				opts.push(`<option value="${val}" ${current == val ? 'selected' : ''}>${val}</option>`);
			}
			return opts.join('');
		};
		const rateSelect = step.type === 'pressure'
			? `<select onchange="updateStepRate(${step.id}, this.value)">${buildRateOptions(0.05, 0.45, 0.05, step.rate || '')}</select>`
			: `<select onchange="updateStepRate(${step.id}, this.value)">${buildRateOptions(0.05, 2.00, 0.05, step.rate || '')}</select>`;
		typeSpecificFields = `
			<div class="form-group">
				<label>速率 (bar/min)</label>
				${rateSelect}
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
	
	// 決定目標壓力輸入的屬性（持壓時鎖定並套用上一段壓力）
	let pressureValue = step.pressure || '';
	let pressureAttrs = 'step="0.1" min="0"';
	if (step.type === 'hold') {
		// 若有上一段壓力，沿用並鎖定；否則保持可輸入
		if (prevStep && prevStep.pressure !== undefined && prevStep.pressure !== '') {
			pressureValue = formatToOneDecimal(prevStep.pressure);
			step.pressure = pressureValue;
			pressureAttrs += ' disabled readonly title="持壓步驟之目標壓力自動沿用上一段，無法手動修改"';
		}
	}
	// 顯示時一律套用一位小數（若非空）
	pressureValue = pressureValue !== '' ? formatToOneDecimal(pressureValue) : '';

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
				<input type="number" ${pressureAttrs} value="${pressureValue}" 
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
			// 將持壓的目標壓力同步為上一段壓力
			const idx = testSteps.findIndex(s => s.id === stepId);
			const prev = idx > 0 ? testSteps[idx - 1] : null;
			if (prev && prev.pressure !== undefined && prev.pressure !== '') {
				step.pressure = formatToOneDecimal(prev.pressure);
			}
		} else {
			delete step.holdTime;
		}
		updateTestSteps();
		autoUpdateTimeline();
	}
}

// 更新步驟速率
function updateStepRate(stepId, rate) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.rate = rate;
		updateTestSteps();
		autoUpdateTimeline();
	}
}

// 更新步驟持壓時間
function updateStepHoldTime(stepId, holdTime) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.holdTime = holdTime;
		updateTestSteps();
		autoUpdateTimeline();
	}
}

// 更新步驟壓力
function updateStepPressure(stepId, pressure) {
	const step = testSteps.find(s => s.id === stepId);
	if (step) {
		step.pressure = formatToOneDecimal(pressure);
		// 若下一步是持壓，則同步其目標壓力
		const idx = testSteps.findIndex(s => s.id === stepId);
		const next = idx >= 0 && idx + 1 < testSteps.length ? testSteps[idx + 1] : null;
		if (next && next.type === 'hold') {
			next.pressure = formatToOneDecimal(pressure);
		}
		updateTestSteps();
		autoUpdateTimeline();
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
			return `增壓至 ${step.pressure} bar，速率 ${step.rate} bar/min。`;
		}
		return `增壓至${step.pressure || ''} bar，速率 ${step.rate || ''} bar/min。`;
	} else if (step.type === 'depressurize') {
		if (step.pressure && step.rate) {
			return `降壓至 ${step.pressure} bar，速率 ${step.rate} bar/min。`;
		}
		return `降壓至${step.pressure || ''} bar，速率 ${step.rate || ''} bar/min。`;
	} else if (step.type === 'hold') {
		if (step.pressure && step.holdTime) {
			return `持壓 ${step.pressure} bar，持續 ${step.holdTime} min。`;
		}
		return `持壓 ${step.pressure || ''} bar，持續 ${step.holdTime || ''} min。`;
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
	autoUpdateTimeline();
}

// 添加預設受測物
function addDefaultTestObject() {
	const defaultObject = {
		id: testObjectCounter++,
		name: '',
		quantity: '1',
		size: '',
		waterStatus: '',
		photos: [],
		notes: ''
	};
	
	testObjects = [defaultObject];
	updateTestObjects();
}

// 添加受測物
function addTestObject() {
	const testObject = {
		id: testObjectCounter++,
		name: '',
		quantity: '1',
		size: '',
		waterStatus: '',
		photos: [],
		notes: ''
	};
	
	testObjects.push(testObject);
	updateTestObjects();
}

// 移除受測物
function removeTestObject(objectId) {
	testObjects = testObjects.filter(obj => obj.id !== objectId);
	updateTestObjects();
}

// 更新受測物顯示
function updateTestObjects() {
	if (!testObjectsList) {
		return;
	}
	
	testObjectsList.innerHTML = '';
	
	testObjects.forEach((obj, index) => {
		const objectDiv = createTestObjectElement(obj, index + 1);
		testObjectsList.appendChild(objectDiv);
	});
}

// 創建受測物元素
function createTestObjectElement(obj, objectNumber) {
	const objectDiv = document.createElement('div');
	objectDiv.className = 'test-object-item';
	objectDiv.innerHTML = `
		<div class="test-object-header">
			<div class="test-object-title">受測物 ${objectNumber}</div>
			<button type="button" class="remove-test-object" onclick="removeTestObject(${obj.id})">移除</button>
		</div>
		<div class="test-object-form">
			<div class="form-group">
				<label>受測物名稱 *</label>
				<input type="text" value="${obj.name || ''}" onchange="updateTestObjectName(${obj.id}, this.value)" required>
			</div>
			<div class="form-group">
				<label>數量 *</label>
				<select onchange="updateTestObjectQuantity(${obj.id}, this.value)" required>
					<option value="">請選擇數量</option>
					<option value="1" ${obj.quantity === '1' ? 'selected' : ''}>1</option>
					<option value="2" ${obj.quantity === '2' ? 'selected' : ''}>2</option>
					<option value="3" ${obj.quantity === '3' ? 'selected' : ''}>3</option>
					<option value="4" ${obj.quantity === '4' ? 'selected' : ''}>4</option>
					<option value="5" ${obj.quantity === '5' ? 'selected' : ''}>5</option>
					<option value="6" ${obj.quantity === '6' ? 'selected' : ''}>6</option>
					<option value="7" ${obj.quantity === '7' ? 'selected' : ''}>7</option>
					<option value="8" ${obj.quantity === '8' ? 'selected' : ''}>8</option>
					<option value="9" ${obj.quantity === '9' ? 'selected' : ''}>9</option>
					<option value="10" ${obj.quantity === '10' ? 'selected' : ''}>10</option>
				</select>
			</div>
			<div class="form-group full-width">
				<label>尺寸 *</label>
				<input type="text" value="${obj.size || ''}" onchange="updateTestObjectSize(${obj.id}, this.value)" required>
			</div>
			<div class="form-group">
				<label>入水狀態 *</label>
				<select onchange="updateTestObjectWaterStatus(${obj.id}, this.value)" required>
					<option value="">請選擇入水狀態</option>
					<option value="下沉" ${obj.waterStatus === '下沉' ? 'selected' : ''}>下沉</option>
					<option value="上浮" ${obj.waterStatus === '上浮' ? 'selected' : ''}>上浮</option>
				</select>
			</div>
			<div class="form-group full-width">
				<label>受測物照片（最多3張）</label>
				<input type="file" accept="image/*" multiple onchange="handleTestObjectPhotosUpload(${obj.id}, this)">
				<div class="test-object-photo-preview" id="photoPreview_${obj.id}" style="display: ${obj.photos && obj.photos.length ? 'block' : 'none'};">
					${(obj.photos || []).map((src, i) => `<div class=\"photo-item\"><img src=\"${src}\" alt=\"受測物照片\"><button type=\"button\" class=\"remove-photo\" onclick=\"removeTestObjectPhoto(${obj.id}, ${i})\">移除</button></div>`).join('')}
				</div>
			</div>
			<div class="form-group full-width">
				<label>備註</label>
				<textarea rows="3" onchange="updateTestObjectNotes(${obj.id}, this.value)">${obj.notes || ''}</textarea>
			</div>
		</div>
	`;
	
	return objectDiv;
}

// 更新受測物名稱
function updateTestObjectName(objectId, name) {
	const obj = testObjects.find(o => o.id === objectId);
	if (obj) {
		obj.name = name;
	}
}

// 更新受測物數量
function updateTestObjectQuantity(objectId, quantity) {
	const obj = testObjects.find(o => o.id === objectId);
	if (obj) {
		obj.quantity = quantity;
	}
}

// 更新受測物尺寸
function updateTestObjectSize(objectId, size) {
	const obj = testObjects.find(o => o.id === objectId);
	if (obj) {
		obj.size = size;
	}
}

// 更新受測物入水狀態
function updateTestObjectWaterStatus(objectId, waterStatus) {
	const obj = testObjects.find(o => o.id === objectId);
	if (obj) {
		obj.waterStatus = waterStatus;
	}
}

// 更新受測物備註
function updateTestObjectNotes(objectId, notes) {
	const obj = testObjects.find(o => o.id === objectId);
	if (obj) {
		obj.notes = notes;
	}
}

// 處理受測物多張照片上傳（最多3張）
function handleTestObjectPhotosUpload(objectId, input) {
	const files = Array.from(input.files || []);
	if (files.length === 0) return;
	const obj = testObjects.find(o => o.id === objectId);
	if (!obj) return;
	if (!Array.isArray(obj.photos)) obj.photos = [];
	const remain = Math.max(0, 3 - obj.photos.length);
	const toRead = files.slice(0, remain);
	if (toRead.length === 0) {
		alert('最多可上傳 3 張照片');
		input.value = '';
		return;
	}
	let pending = toRead.length;
	toRead.forEach(file => {
		const reader = new FileReader();
		reader.onload = function(e) {
			obj.photos.push(e.target.result);
			pending -= 1;
			if (pending === 0) {
				updateTestObjects();
			}
		};
		reader.readAsDataURL(file);
	});
	input.value = '';
}

// 移除指定索引的照片
function removeTestObjectPhoto(objectId, index) {
	const obj = testObjects.find(o => o.id === objectId);
	if (!obj || !Array.isArray(obj.photos)) return;
	obj.photos.splice(index, 1);
	updateTestObjects();
}

// 移除不再使用的照片上傳函數
// function handlePhotoUpload(event) {
// 	const file = event.target.files[0];
// 	if (file) {
// 		const reader = new FileReader();
// 		reader.onload = function(e) {
// 			photoPreview.innerHTML = `<img src="${e.target.result}" alt="設備照片">`;
// 			photoPreview.style.display = 'block';
// 		};
// 		reader.readAsDataURL(file);
// 	}
// }

// 自動更新時序圖
function autoUpdateTimeline() {
	// 檢查是否有有效的測試步驟
	const validSteps = testSteps.filter(step => {
		if (step.type === 'pressure' || step.type === 'depressurize') {
			return step.pressure && step.rate;
		} else if (step.type === 'hold') {
			return step.pressure && step.holdTime;
		}
		return false;
	});
	
	// 如果有有效步驟，自動繪製時序圖
	if (validSteps.length > 0) {
		drawTimeline();
		// 顯示時序圖容器
		document.getElementById('timelineContainer').style.display = 'block';
	} else {
		// 如果沒有有效步驟，隱藏時序圖容器
		document.getElementById('timelineContainer').style.display = 'none';
	}
}

// 生成時序圖（保留原函數，但簡化邏輯）
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
	try {
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
		testObjects: testObjects,
		testSteps: testSteps,
		wiring: collectWiringData()
	};
	
	const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, ImageRun, BorderStyle } = window.docx;
	
	// 標題
	const title = new Paragraph({
		children: [
			new TextRun({
				text: '高壓模擬實驗觀測系統使用申請表',
				bold: true,
				size: 40,
				font: '微軟正黑體'
			})
		],
		heading: HeadingLevel.TITLE,
		alignment: AlignmentType.CENTER,
		spacing: { before: 200, after: 400 }
	});
	
	// 申請日期
	const dateParagraph = new Paragraph({
		children: [
			new TextRun({
				text: `申請日期：${new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: 'long', day: 'numeric' })}`,
				font: '微軟正黑體',
				size: 24
			})
		],
		alignment: AlignmentType.RIGHT,
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
			new TableCell({ 
				children: [new Paragraph({
					children: [new TextRun({ text: k, bold: true, font: '微軟正黑體', size: 24 })],
					alignment: AlignmentType.CENTER
				})],
				width: { size: 25, type: WidthType.PERCENTAGE },
				shading: { fill: 'F2F2F2' }
			}),
			new TableCell({ 
				children: [new Paragraph({
					children: [new TextRun({ text: v, font: '微軟正黑體', size: 24 })],
					alignment: AlignmentType.CENTER
				})],
				width: { size: 75, type: WidthType.PERCENTAGE }
			})
		]
	}));
	const companyTable = new Table({
		rows: companyRows,
		width: { size: 100, type: WidthType.PERCENTAGE },
		borders: {
			top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' }
		}
	});
	
	// 受測物資料表格
	let testObjectsContent = [];
	
	if (formData.testObjects && formData.testObjects.length > 0) {
		formData.testObjects.forEach((obj, index) => {
			testObjectsContent.push(
				new Paragraph({
					children: [new TextRun({
						text: `受測物 ${index + 1}`,
						font: '微軟正黑體',
						size: 28
					})],
					heading: HeadingLevel.HEADING_3,
					spacing: { before: 200, after: 100 }
				})
			);
			
			const objectRows = [
				['受測物名稱', obj.name || ''],
				['數量', obj.quantity || ''],
				['尺寸', obj.size || ''],
				['入水狀態', obj.waterStatus || ''],
				['備註', obj.notes || '']
			].map(([k, v]) => new TableRow({
				children: [
					new TableCell({ 
						children: [new Paragraph({
							children: [new TextRun({ text: k, bold: true, font: '微軵正黑體', size: 24 })],
							alignment: AlignmentType.CENTER
						})],
						width: { size: 25, type: WidthType.PERCENTAGE },
						shading: { fill: 'F2F2F2' }
					}),
					new TableCell({ 
						children: [new Paragraph({
							children: [new TextRun({ text: v, font: '微軟正黑體', size: 24 })],
							alignment: AlignmentType.CENTER
						})],
						width: { size: 75, type: WidthType.PERCENTAGE }
					})
				]
			}));
			
			const objectTable = new Table({
				rows: objectRows,
				width: { size: 100, type: WidthType.PERCENTAGE },
				borders: {
					top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
					bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
					left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
					right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
					insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
					insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' }
				}
			});
			
			testObjectsContent.push(objectTable);
			
			// 如果有照片（多張），添加照片
			if (obj.photos && Array.isArray(obj.photos) && obj.photos.length > 0) {
				try {
					testObjectsContent.push(new Paragraph({
						text: '受測物照片：',
						spacing: { before: 200, after: 100 }
					}));
					obj.photos.forEach((photoDataUrl) => {
						try {
							const imageData = photoDataUrl.split(',')[1];
							const imageBuffer = Uint8Array.from(atob(imageData), c => c.charCodeAt(0));
							testObjectsContent.push(new Paragraph({
								children: [
									new ImageRun({
										data: imageBuffer,
										transformation: { width: 300, height: 200 },
									}),
								],
								alignment: AlignmentType.CENTER,
								espacing: { after: 100 }
							}));
						} catch (innerErr) {
							console.error('處理單張照片時發生錯誤:', innerErr);
						}
					});
				} catch (error) {
					console.error('處理照片時發生錯誤:', error);
				}
			}
		});
	} else {
		testObjectsContent.push(
			new Paragraph({
				text: '無受測物資料',
				spacing: { before: 200, after: 100 }
			})
		);
	}
	
	// 測試要求（編號清單）
	const stepParas = [];
	formData.testSteps.forEach((s, idx) => {
		const desc = s.description && s.description.trim().length > 0 ? s.description : generateStepDescription(s);
		stepParas.push(new Paragraph({ 
			children: [new TextRun({ 
				text: `${idx + 1}. ${desc}`, 
				font: '微軟正黑體',
				size: 24
			})], 
			spacing: { after: 120 } 
		}));
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
	const h1 = (t) => new Paragraph({ 
		heading: HeadingLevel.HEADING_1, 
		spacing: { before: 400, after: 200 },
		children: [
			new TextRun({
				text: t,
				bold: true,
				size: 32,
				color: '2C3E50',
				font: '微軟正黑體'
			})
		]
	});
	
	// 接線要求內容
	const wiringParas = [];
	if (formData.wiring.selectedSystem) {
		// 不再輸出「選用系統規格說明」與系統規格敘述
		
		if (formData.wiring.selectedSystem === '450') {
			if (formData.wiring.system450.length > 0) {
				wiringParas.push(new Paragraph({ 
					children: [new TextRun({
						text: 'HPT-450-230L 系統接線要求：',
						font: '微軟正黑體',
						size: 24
					})],
					spacing: { before: 200, after: 100 } 
				}));
				formData.wiring.system450.forEach(item => {
					const componentText = item.componentSpec ? ` - 接頭零件規格：${item.componentSpec}` : '';
					const remarkText = item.remark ? ` - 備註：${item.remark}` : '';
					wiringParas.push(new Paragraph({ 
						children: [new TextRun({
							text: `(${item.port}) ${item.spec}${item.description ? ': ' + item.description : ''}${componentText}${remarkText}`,
							font: '微軟正黑體',
							size: 24
						})],
						spacing: { after: 80 } 
					}));
				});
			}
		} else if (formData.wiring.selectedSystem === '800') {
			if (formData.wiring.system800.length > 0) {
				wiringParas.push(new Paragraph({ 
					children: [new TextRun({
						text: 'HPT-800-85L 系統接線要求：',
						font: '微軟正黑體',
						size: 24
					})],
					spacing: { before: 200, after: 100 } 
				}));
				formData.wiring.system800.forEach(item => {
					const componentText = item.componentSpec ? ` - 接頭零件規格：${item.componentSpec}` : '';
					const remarkText = item.remark ? ` - 備註：${item.remark}` : '';
					wiringParas.push(new Paragraph({ 
						children: [new TextRun({
							text: `(${item.port}) ${item.spec}${item.description ? ': ' + item.description : ''}${componentText}${remarkText}`,
							font: '微軟正黑體',
							size: 24
						})],
						spacing: { after: 80 } 
					}));
				});
			}
		}
	}
	
	// 建立文件
	const docChildren = [
		title,
		dateParagraph,
		h1('一、委託單位資料'),
		companyTable,
		h1('二、受測物資料'),
		...testObjectsContent,
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
	
	// 添加價格試算
	try {
		const pricingContent = generatePricingContent();
		if (pricingContent.length > 0) {
			docChildren.push(h1('六、價格試算'));
			docChildren.push(...pricingContent);
		}
	} catch (error) {
		console.error('生成價格試算內容時發生錯誤:', error);
		// 跳過價格試算，繼續生成文件
	}

	const doc = new Document({
		sections: [{ children: docChildren }]
	});
	
		const blob = await Packer.toBlob(doc);
		const fileName = `高壓艙申請表_${new Date().toISOString().slice(0,10)}.docx`;
		triggerDownload(blob, fileName);
	} catch (error) {
		console.error('生成Word文件時發生錯誤:', error);
		alert('生成文件時發生錯誤，請檢查控制台以獲取詳細信息。');
	}
}

function dataURLToUint8Array(dataURL) {
	const base64 = dataURL.split(',')[1];
	const binary = atob(base64);
	const len = binary.length;
	const bytes = new Uint8Array(len);
	for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
	return bytes;
}

// 生成價格試算內容
function generatePricingContent() {
	const { Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel, BorderStyle, TextRun } = window.docx;
	
	// 獲取測試參數
	const testParams = getTestParameters();
	if (!testParams || testParams.totalTime === 0) {
		return [];
	}
	
	// 計算價格
	const pricing = calculateBasePricing(testParams.maxPressure, Math.round(testParams.totalTime));
	const additionalFees = calculateAdditionalFees(pricing.basicTestFee);
	
	// 獲取額外服務選擇
	const englishReportCheckbox = document.getElementById('englishReport');
	const videoRecordingCheckbox = document.getElementById('videoRecording');
	const englishReport = englishReportCheckbox && englishReportCheckbox.checked ? additionalFees.englishReport : 0;
	const videoRecording = videoRecordingCheckbox && videoRecordingCheckbox.checked ? additionalFees.videoRecording : 0;
	
	const total = pricing.basicTestFee + englishReport + videoRecording;
	
	// 轉換時間格式
	const formatTime = (minutes) => {
		const hours = Math.floor(minutes / 60);
		const mins = minutes % 60;
		return mins > 0 ? `${hours} hr ${mins} min` : `${hours} hr`;
	};
	
	const content = [];
	
	// 測試參數
	content.push(new Paragraph({
		children: [new TextRun({
			text: '測試參數',
			font: '微軟正黑體',
			size: 28
		})],
		heading: HeadingLevel.HEADING_3,
		spacing: { before: 200, after: 100 }
	}));
	
	const paramsRows = [
		['最大壓力', `${testParams.maxPressure} bar`],
		['總測試時間', formatTime(testParams.totalTime)],
		['選擇系統', testParams.systemName || '未選擇']
	].map(([k, v]) => new TableRow({
		children: [
			new TableCell({ 
				children: [new Paragraph({
					children: [new TextRun({ text: k, bold: true, font: '微軟正黑體', size: 24 })],
					alignment: AlignmentType.CENTER
				})],
				width: { size: 25, type: WidthType.PERCENTAGE },
				shading: { fill: 'F2F2F2' }
			}),
			new TableCell({ 
				children: [new Paragraph({
					children: [new TextRun({ text: v, font: '微軟正黑體', size: 24 })],
					alignment: AlignmentType.CENTER
				})],
				width: { size: 75, type: WidthType.PERCENTAGE }
			})
		]
	}));
	
	content.push(new Table({
		rows: paramsRows,
		width: { size: 100, type: WidthType.PERCENTAGE },
		borders: {
			top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' }
		}
	}));
	
	// 價格計算
	content.push(new Paragraph({
		children: [new TextRun({
			text: '價格計算',
			font: '微軟正黑體',
			size: 28
		})],
		heading: HeadingLevel.HEADING_3,
		spacing: { before: 200, after: 100 }
	}));
	
	const pricingRows = [
		['基本費用 (3小時內)', `NT$ ${pricing.baseFee.toLocaleString()}`]
	];
	
	if (pricing.overtimeFee > 0) {
		pricingRows.push(['超時費用 (第4-8小時)', `NT$ ${pricing.overtimeFee.toLocaleString()}`]);
		pricingRows.push(['超時小時數', formatTime(pricing.overtimeMinutes)]);
	}
	
	if (pricing.extendedOvertimeFee > 0) {
		pricingRows.push(['延長超時費用 (第9小時起)', `NT$ ${pricing.extendedOvertimeFee.toLocaleString()}`]);
		pricingRows.push(['延長超時小時數', formatTime(pricing.extendedOvertimeMinutes)]);
	}
	
	pricingRows.push(['基本測試費用小計', `NT$ ${pricing.basicTestFee.toLocaleString()}`]);
	
	if (englishReport > 0) {
		pricingRows.push(['英文測試報告', `NT$ ${englishReport.toLocaleString()}`]);
	}
	
	if (videoRecording > 0) {
		pricingRows.push(['錄影服務', `NT$ ${videoRecording.toLocaleString()}`]);
	}
	
	pricingRows.push(['總計', `NT$ ${total.toLocaleString()}`]);
	
	const pricingTableRows = pricingRows.map(([k, v]) => new TableRow({
		children: [
			new TableCell({ 
				children: [new Paragraph({
					children: [new TextRun({ text: k, bold: true, font: '微軟正黑體', size: 24 })],
					alignment: AlignmentType.CENTER
				})],
				width: { size: 25, type: WidthType.PERCENTAGE },
				shading: { fill: k === '總計' ? 'E8F4FD' : 'F2F2F2' }
			}),
			new TableCell({ 
				children: [new Paragraph({
					children: [new TextRun({ 
						text: v, 
						bold: k === '總計',
						color: k === '總計' ? '2C3E50' : '000000',
						font: '微軟正黑體',
						size: 24
					})],
					alignment: AlignmentType.CENTER
				})],
				width: { size: 75, type: WidthType.PERCENTAGE },
				shading: { fill: k === '總計' ? 'E8F4FD' : 'FFFFFF' }
			})
		]
	}));
	
	content.push(new Table({
		rows: pricingTableRows,
		width: { size: 100, type: WidthType.PERCENTAGE },
		borders: {
			top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
			insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' }
		}
	}));
	
	// 免責聲明
	content.push(new Paragraph({
		children: [
			new TextRun({
				text: '※ 以上價格為試算價格，實際價格以最終報價為準。',
				italic: true,
				color: '666666',
				size: 24,
				font: '微軟正黑體'
			})
		],
		spacing: { before: 300, after: 200 },
		alignment: AlignmentType.CENTER
	}));
	
	return content;
}

function triggerDownload(blob, fileName) {
	// 強制設定正確的 MIME 類型，避免雲端/瀏覽器誤判
	const docBlob = new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
	const url = URL.createObjectURL(docBlob);
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

// 初始化接線要求
function initializeWiring() {
	// 設置系統選擇事件
	setupSystemSelection();
	
	// 設置接線選項事件
	setupWiringCheckboxes();
	
	// 初始化預設選中的系統（HPT-450-230L）
	const system450Radio = document.getElementById('system450');
	if (system450Radio) {
		// 觸發change事件來顯示對應的系統選項
		system450Radio.dispatchEvent(new Event('change'));
	}
	// 初始化時依最大壓力限制可選系統
	enforceSystemEligibility();
}

// 依最大壓力限制系統選擇（>450 bar 禁用 HPT-450-230L）
function enforceSystemEligibility() {
	const params = getTestParameters();
	const radio450 = document.getElementById('system450');
	const radio800 = document.getElementById('system800');
	const label450 = radio450 ? radio450.closest('.system-radio-label') : null;
	if (!radio450 || !radio800) return;
	const maxP = parseFloat(params && params.maxPressure ? params.maxPressure : 0);
	if (!isNaN(maxP) && maxP > 450) {
		radio450.disabled = true;
		if (label450) label450.classList.add('disabled');
		if (radio450.checked) {
			radio450.checked = false;
			radio800.checked = true;
			radio800.dispatchEvent(new Event('change'));
		}
		radio450.title = '目標壓力超過 450 bar，無法選用 HPT-450-230L';
	} else {
		radio450.disabled = false;
		if (label450) label450.classList.remove('disabled');
		radio450.title = '';
	}
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
		updateTestParamsDisplay('-- bar', '-- hr', '--');
		updatePricingDisplay(0, 0, 0, 0, 0, 0);
		return;
	}
	
	// 更新測試參數顯示
	const totalMinutes = Math.round(testParams.totalTime);
	const hours = Math.floor(totalMinutes / 60);
	const minutes = totalMinutes % 60;
	const timeDisplay = minutes > 0 ? `${hours} hr ${minutes} min` : `${hours} hr`;
	updateTestParamsDisplay(
		`${testParams.maxPressure} bar`,
		timeDisplay,
		testParams.systemName
	);

	// 控制「不足一小時以一小時計算」說明顯示條件：>=3小時且有分鐘
	const priceNoteEl = document.getElementById('priceNote');
	if (priceNoteEl) {
		if (totalMinutes >= 180 && minutes > 0) {
			priceNoteEl.style.display = 'block';
		} else {
			priceNoteEl.style.display = 'none';
		}
	}

	// 計算基本費用（直接使用分鐘）
	const pricing = calculateBasePricing(testParams.maxPressure, totalMinutes);
	
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

	// 更新標準測試費用小計（基本費 + 超時 + 延長）
	const standardSubtotalEl = document.getElementById('standardSubtotal');
	if (standardSubtotalEl) {
		standardSubtotalEl.textContent = `NT$ ${pricing.basicTestFee.toLocaleString()}`;
	}
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
		// 第4-8小時（最多5小時 = 300分鐘）
		const regularOvertimeMinutes = Math.min(totalTimeMinutes - threeHoursInMinutes, 5 * 60);
		overtimeMinutes = regularOvertimeMinutes;
		const regularChargeHours = Math.ceil(regularOvertimeMinutes / 60);
		overtimeFee = regularChargeHours * overtimeRatePerHour;
		
		// 第9小時起（1.5倍）
		if (totalTimeMinutes > eightHoursInMinutes) {
			extendedOvertimeMinutes = totalTimeMinutes - eightHoursInMinutes;
			const extendedChargeHours = Math.ceil(extendedOvertimeMinutes / 60);
			extendedOvertimeFee = Math.round(extendedChargeHours * overtimeRatePerHour * 1.5);
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
		const roundedOvertimeH = Math.ceil(overtimeMinutes / 60);
		const overtimeDisplay = overtimeM > 0 ? `${overtimeH} hr ${overtimeM} min (計${roundedOvertimeH} hr)` : `${overtimeH} hr`;
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
		const roundedExtendedH = Math.ceil(extendedOvertimeMinutes / 60);
		const extendedDisplay = extendedM > 0 ? `${extendedH} hr ${extendedM} min (計${roundedExtendedH} hr)` : `${extendedH} hr`;
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

// 第2階段必填檢查
function validatePage2Required() {
	const name = document.getElementById('companyName')?.value?.trim();
	const taxId = document.getElementById('taxId')?.value?.trim();
	const address = document.getElementById('address')?.value?.trim();
	const contactPerson = document.getElementById('contactPerson')?.value?.trim();
	const contactPhone = document.getElementById('contactPhone')?.value?.trim();
	const contactEmail = document.getElementById('contactEmail')?.value?.trim();
	return !!(name && taxId && address && contactPerson && contactPhone && contactEmail);
}

// 第3階段必填檢查（依據 testObjects 資料）
function validatePage3Required() {
	if (!Array.isArray(testObjects) || testObjects.length === 0) return false;
	for (let i = 0; i < testObjects.length; i++) {
		const obj = testObjects[i];
		if (!obj) return false;
		const hasAll = (obj.name && obj.name.trim()) && (obj.quantity && obj.quantity.trim()) && (obj.size && obj.size.trim()) && (obj.waterStatus && obj.waterStatus.trim());
		if (!hasAll) return false;
	}
	return true;
}
