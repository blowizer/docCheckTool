// 全局缓存比对结果，便于同步显示
// let lastCompareResult = null;

document.getElementById("requirementFile").addEventListener("change", function () {
    document.getElementById("requirementFileName").value = this.files[0]?.name || "";
});

document.getElementById("designFile").addEventListener("change", function () {
    document.getElementById("designFileName").value = this.files[0]?.name || "";
});

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
            const arrayBuffer = e.target.result;

            // mammoth.js 转换选项
            const options = {
                arrayBuffer: arrayBuffer,
                // 可以添加转换器来过滤特定内容
                transformDocument: function (element) {
                    // 过滤掉嵌入式对象（如压缩文件）
                    element.children = element.children.filter(function (child) {
                        return child.type !== "embedded";
                    });
                    return element;
                }
            };

            mammoth.convertToHtml(options)
                .then(result => resolve(result.value))
                .catch(error => reject(error));
        };
        reader.onerror = () => reject(new Error("文件读取失败"));
        reader.readAsArrayBuffer(file);
    });
}

function extractInfo(html, selector) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const elements = doc.querySelectorAll(selector);
    // 先移除所有嵌入式对象
    const embeddedObjects = doc.querySelectorAll('object, embed');
    embeddedObjects.forEach(obj => obj.remove());
    const info = {};

    // 只处理li元素，直接提取所有功能编号和名称
    if (selector === 'li') {
        elements.forEach((element) => {
            const text = element.textContent.trim();
            if (/功能编号[：:]/.test(text)) {
                const functionId = text.split(/[:：]/).pop().trim();
                const functionNameElement = element.nextElementSibling;
                if (functionNameElement) {
                    const functionName = functionNameElement.textContent.split(/[:：]/).pop().trim();
                    info[functionId] = functionName;
                }
            }
        });
    } else if (selector === 'td') {
        // 只提取最后一个table中的功能编号和名称
        const tables = doc.querySelectorAll('table');
        if (tables.length > 0) {
            const lastTable = tables[tables.length - 1];
            const rows = lastTable.querySelectorAll('tr');
            rows.forEach((row) => {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 2) {
                    // 假设第一列为功能编号，第二列为功能名称
                    const functionId = cells[0].textContent.trim();
                    const functionName = cells[1].textContent.trim();
                    // 跳过表头或key为"功能编号"的行
                    if (functionId === '功能编号') return;
                    if (functionId && functionName) {
                        info[functionId] = functionName;
                    }
                }
            });
        }
    }

    return info;
}

function compareInfo(requirementInfo, designInfo) {
    const errors = [];
    for (const [functionId, functionName] of Object.entries(requirementInfo)) {
        if (!designInfo.hasOwnProperty(functionId)) {
            errors.push(`功能编号 ${functionId} 在概要设计文档中未找到。`);
        } else if (designInfo[functionId] !== functionName) {
            errors.push(`功能编号 ${functionId} 的功能描述不一致：需求文档为 ${functionName}，概要设计文档为 ${designInfo[functionId]}。`);
        }
    }
    for (const functionId in designInfo) {
        if (!requirementInfo.hasOwnProperty(functionId)) {
            errors.push(`功能编号 ${functionId} 在需求文档中未找到。`);
        }
    }
    return errors;
}

function displayResult(errors, requirementInfo, designInfo) {
    let resultDiv = document.getElementById("result");
    if (!resultDiv) {
        console.log("找不到 result div 元素，正在创建...");
        // 创建 result div 并插入到正确位置
        resultDiv = document.createElement('div');
        resultDiv.id = 'result';

        // 找到功能编号检查按钮，在其后插入 result div
        const checkFuncMatchBtn = document.getElementById('checkFuncMatchBtn');
        if (checkFuncMatchBtn && checkFuncMatchBtn.parentNode) {
            checkFuncMatchBtn.parentNode.insertBefore(resultDiv, checkFuncMatchBtn.nextSibling);
        } else {
            // 如果找不到按钮，插入到检查表单的末尾
            const checklistForm = document.getElementById('checklistForm');
            if (checklistForm) {
                checklistForm.appendChild(resultDiv);
            } else {
                console.error("无法找到合适的位置插入 result div");
                alert("显示结果时出错，请刷新页面重试");
                return;
            }
        }
    }

    let html = '';
    // 输出比对结果（最上方）
    if (errors.length === 0) {
        html += "<div class='error-details'><div class='error-title' style='color:green;'>比对成功！</div></div>";
    } else {
        html += `<div class='error-details'><div class='error-title'>比对结果：</div><ul style='margin:0;padding-left:18px;'>${errors.map(e => `<li>${e}</li>`).join("")}</ul></div>`;
    }
    // 输出需求文档功能编号和名称
    html += '<div class="error-details">';
    html += '<div class="error-title">需求文档功能列表：</div>';
    html += '<div class="error-item">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr><th style="text-align:left;padding:2px 6px;">功能编号</th><th style="text-align:left;padding:2px 6px;">功能名称</th></tr>';
    for (const [id, name] of Object.entries(requirementInfo)) {
        html += `<tr><td style="padding:2px 6px;border-bottom:1px solid #eee;">${id}</td><td style="padding:2px 6px;border-bottom:1px solid #eee;">${name}</td></tr>`;
    }
    html += '</table>';
    html += '</div>';
    html += '</div>';
    // 输出概要设计文档功能编号和名称
    html += '<div class="error-details">';
    html += '<div class="error-title">概要设计文档功能列表：</div>';
    html += '<div class="error-item">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr><th style="text-align:left;padding:2px 6px;">功能编号</th><th style="text-align:left;padding:2px 6px;">功能名称</th></tr>';
    for (const [id, name] of Object.entries(designInfo)) {
        html += `<tr><td style="padding:2px 6px;border-bottom:1px solid #eee;">${id}</td><td style="padding:2px 6px;border-bottom:1px solid #eee;">${name}</td></tr>`;
    }
    html += '</table>';
    html += '</div>';
    html += '</div>';

    try {
        resultDiv.innerHTML = html;
    } catch (error) {
        console.error("设置 innerHTML 时出错:", error);
        alert("显示结果时出错，请刷新页面重试");
    }
    // 不再缓存比对结果
}

// 页签切换功能
window.addEventListener('DOMContentLoaded', function () {
    // 页签切换
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', function () {
            const targetTab = this.getAttribute('data-tab');

            // 移除所有活动状态
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));

            // 添加当前活动状态
            this.classList.add('active');
            document.getElementById(targetTab + '-tab').classList.add('active');
        });
    });

    // CR名称检查按钮
    const checkCRBtn = document.getElementById('checkCRBtn');
    if (checkCRBtn) {
        checkCRBtn.addEventListener('click', function () {
            checkCRName();
        });
    }

    // 需求标题检查按钮
    const checkTitleBtn = document.getElementById('checkTitleBtn');
    if (checkTitleBtn) {
        checkTitleBtn.addEventListener('click', function () {
            checkRequirementTitle();
        });
    }

    // 功能编号和名称是否匹配 检查按钮
    const checkFuncMatchBtn = document.getElementById('checkFuncMatchBtn');
    if (checkFuncMatchBtn) {
        checkFuncMatchBtn.addEventListener('click', function () {
            checkFuncMatch();
        });
    }

    // 上传页签"去检查"按钮
    const gotoCheckBtn = document.getElementById('gotoCheckBtn');
    if (gotoCheckBtn) {
        gotoCheckBtn.addEventListener('click', function () {
            // 切换tab
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelector('.tab-btn[data-tab="check"]').classList.add('active');
            document.getElementById('check-tab').classList.add('active');
        });
    }

    // 功能模块检查按钮
    const checkModuleBtn = document.getElementById('checkModuleBtn');
    if (checkModuleBtn) {
        checkModuleBtn.addEventListener('click', function () {
            checkModuleMatch();
        });
    }

    // 人员检查按钮
    const checkPersonBtn = document.getElementById('checkPersonBtn');
    if (checkPersonBtn) {
        checkPersonBtn.addEventListener('click', checkPersonMatch);
    }

    // 接口需求检查按钮
    const checkInterfaceBtn = document.getElementById('checkInterfaceBtn');
    if (checkInterfaceBtn) {
        checkInterfaceBtn.addEventListener('click', checkInterfaceRequirement);
    }

    // 概要设计文档接口检查按钮
    const checkDesignInterfaceBtn = document.getElementById('checkDesignInterfaceBtn');
    if (checkDesignInterfaceBtn) {
        checkDesignInterfaceBtn.addEventListener('click', checkDesignInterfaceRequirement);
    }

    // 一键检查按钮
    const oneClickCheckBtn = document.getElementById('oneClickCheckBtn');
    if (oneClickCheckBtn) {
        oneClickCheckBtn.addEventListener('click', performOneClickCheck);
    }

    // 导出结果按钮
    const exportResultBtn = document.getElementById('exportResultBtn');
    if (exportResultBtn) {
        exportResultBtn.addEventListener('click', exportCheckResults);
    }

    // 导出格式选择器
    const exportFormatSelect = document.getElementById('exportFormat');
    if (exportFormatSelect) {
        exportFormatSelect.addEventListener('change', function () {
            if (this.value) {
                // 添加导出中状态
                this.classList.add('exporting');
                this.disabled = true;

                // 延迟一下让用户看到加载动画
                setTimeout(() => {
                    exportCheckResults(this.value);

                    // 显示成功状态
                    this.classList.remove('exporting');
                    this.classList.add('success');

                    // 2秒后恢复状态
                    setTimeout(() => {
                        this.classList.remove('success');
                        this.disabled = false;
                        this.value = '';
                    }, 2000);
                }, 500);
            }
        });
    }
});

// CR名称检查功能
function checkCRName() {
    const crNameInput = document.getElementById('crNameInput');
    const checkBtn = document.getElementById('checkCRBtn');
    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];

    if (!requirementFile || !designFile) {
        alert('请先上传需求文档和概要设计文档');
        return;
    }

    if (!crNameInput.value.trim()) {
        alert('请输入CR编号');
        return;
    }

    // 禁用按钮防止重复点击
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';

    const crNumber = crNameInput.value.trim();
    const expectedRequirementName = `${crNumber}_远程银行系统_需求规格说明书.docx`;
    const expectedDesignName = `${crNumber}_远程银行系统_概要设计说明书.docx`;

    // 检查文件名是否匹配
    const requirementMatch = requirementFile.name === expectedRequirementName;
    const designMatch = designFile.name === expectedDesignName;
    const isAllMatch = requirementMatch && designMatch;

    // 更新检查项状态
    updateChecklistStatus('docName', isAllMatch);

    // 显示错误详情
    if (!isAllMatch) {
        showErrorDetails('docName', requirementMatch, designMatch, expectedRequirementName, expectedDesignName, requirementFile.name, designFile.name);
    } else {
        hideErrorDetails('docName');
    }

    // 恢复按钮状态
    checkBtn.disabled = false;
    checkBtn.textContent = '检查';
}

// 提取文档标题
function extractDocumentTitle(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    // 尝试多种方式提取标题
    let title = '';

    // 1. 查找h1标签
    const h1 = doc.querySelector('h1');
    if (h1) {
        title = h1.textContent.trim();
    }

    // 2. 如果没有h1，查找第一个段落
    if (!title) {
        const firstP = doc.querySelector('p');
        if (firstP) {
            title = firstP.textContent.trim();
        }
    }

    // 3. 如果还是没有，查找包含"需求"的文本
    if (!title) {
        const allText = doc.body.textContent;
        const lines = allText.split('\n').map(line => line.trim()).filter(line => line);
        for (let line of lines) {
            if (line.includes('需求') || line.includes('CR')) {
                title = line;
                break;
            }
        }
    }

    return title;
}

// 更新检查项状态
function updateChecklistStatus(itemName, isPass) {
    const checklistItem = document.querySelector(`input[name="${itemName}"]`).closest('.checklist-item');

    // 移除现有的状态图标
    const existingIcon = checklistItem.querySelector('.status-icon');
    if (existingIcon) {
        existingIcon.remove();
    }

    // 创建新的状态图标
    const statusIcon = document.createElement('span');
    statusIcon.className = `status-icon ${isPass ? 'success' : 'error'}`;
    statusIcon.textContent = isPass ? '✓' : '✗';

    // 添加到检查项
    checklistItem.appendChild(statusIcon);

    // 自动选择对应的单选按钮
    const radioBtn = checklistItem.querySelector(`input[name="${itemName}"][value="${isPass ? 'yes' : 'no'}"]`);
    if (radioBtn) {
        radioBtn.checked = true;
    }
}

// 需求标题检查功能
function checkRequirementTitle() {
    const titleInput = document.getElementById('requirementTitleInput');
    const checkBtn = document.getElementById('checkTitleBtn');
    const requirementFile = document.getElementById('requirementFile').files[0];

    if (!requirementFile) {
        alert('请先上传需求文档');
        return;
    }

    if (!titleInput.value.trim()) {
        alert('请输入需求标题');
        return;
    }

    // 禁用按钮防止重复点击
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';

    readFile(requirementFile)
        .then(requirementHtml => {
            // 提取第二个<p>元素内容
            const parser = new DOMParser();
            const doc = parser.parseFromString(requirementHtml, 'text/html');
            let docTitle = '';
            const pList = doc.querySelectorAll('p');
            if (pList.length >= 2) {
                docTitle = pList[1].textContent.trim();
            }
            const inputTitle = titleInput.value.trim() + '_远程银行系统';
            // 比对
            const isMatch = docTitle === inputTitle;
            // 更新检查项状态
            updateChecklistStatus('requirementTitle', isMatch);
            // 总是显示结果，无论通过还是不通过
            showTitleErrorDetails('requirementTitle', docTitle, inputTitle);
            // 恢复按钮状态
            checkBtn.disabled = false;
            checkBtn.textContent = '检查';
        })
        .catch(error => {
            console.error('检查需求标题时出错:', error);
            alert('检查失败，请确保文档格式正确');
            checkBtn.disabled = false;
            checkBtn.textContent = '检查';
        });
}

// 比对需求标题
function compareRequirementTitle(docTitle, inputTitle) {
    if (!docTitle || !inputTitle) return false;

    // 清理和标准化文本
    const cleanDocTitle = docTitle.replace(/[^\w\u4e00-\u9fa5]/g, '').toLowerCase();
    const cleanInputTitle = inputTitle.replace(/[^\w\u4e00-\u9fa5]/g, '').toLowerCase();

    // 检查是否包含关系
    return cleanDocTitle.includes(cleanInputTitle) || cleanInputTitle.includes(cleanDocTitle);
}

// 显示错误详情
function showErrorDetails(itemName, requirementMatch, designMatch, expectedRequirementName, expectedDesignName, actualRequirementName, actualDesignName) {
    // 移除现有的错误详情
    hideErrorDetails(itemName);

    const checklistItem = document.querySelector(`input[name="${itemName}"]`).closest('.checklist-item');
    if (!checklistItem) {
        console.error(`找不到检查项: ${itemName}`);
        return;
    }

    const errorDetails = document.createElement('div');
    errorDetails.className = 'error-details';
    errorDetails.id = `${itemName}-error-details`;

    let errorContent = '';
    if (requirementMatch && designMatch) {
        // 通过时显示绿色成功信息
        errorContent = '<div class="error-title" style="color:green;">✓ 文档名称检查通过！</div>';
    } else {
        // 不通过时显示错误详情
        errorContent = '<div class="error-title">文档名称不匹配详情：</div>';

        if (!requirementMatch) {
            errorContent += `<div class="error-item">
                <span class="error-label">需求文档：</span><br>
                <span class="error-expected">期望：${expectedRequirementName}</span><br>
                <span class="error-actual">实际：${actualRequirementName}</span>
            </div>`;
        }

        if (!designMatch) {
            errorContent += `<div class="error-item">
                <span class="error-label">概要设计文档：</span><br>
                <span class="error-expected">期望：${expectedDesignName}</span><br>
                <span class="error-actual">实际：${actualDesignName}</span>
            </div>`;
        }
    }

    errorDetails.innerHTML = errorContent;

    // 确保插入到正确位置
    if (checklistItem.parentNode) {
        checklistItem.parentNode.insertBefore(errorDetails, checklistItem.nextSibling);
    } else {
        console.error(`检查项 ${itemName} 的父节点不存在`);
    }
}

// 隐藏错误详情
function hideErrorDetails(itemName) {
    const existingError = document.getElementById(`${itemName}-error-details`);
    if (existingError) {
        existingError.remove();
    }
}

// 显示标题错误详情
function showTitleErrorDetails(itemName, docTitle, inputTitle) {
    // 移除现有的标题错误详情
    hideTitleErrorDetails(itemName);

    const checklistItem = document.querySelector(`input[name="${itemName}"]`).closest('.checklist-item');
    if (!checklistItem) {
        console.error(`找不到检查项: ${itemName}`);
        return;
    }

    const errorDetails = document.createElement('div');
    errorDetails.className = 'error-details';
    errorDetails.id = `${itemName}-error-details`;

    let errorContent = '';
    if (docTitle === inputTitle) {
        // 通过时显示绿色成功信息
        errorContent = '<div class="error-title" style="color:green;">✓ 需求标题检查通过！</div>';
    } else {
        // 不通过时显示错误详情
        errorContent = '<div class="error-title">需求标题不匹配详情：</div>';

        errorContent += `<div class="error-item">
            <span class="error-label">文档标题：</span><br>
            <span class="error-expected">期望：${docTitle}</span><br>
            <span class="error-actual">实际：${inputTitle}</span>
        </div>`;
    }

    errorDetails.innerHTML = errorContent;

    // 确保插入到正确位置
    if (checklistItem.parentNode) {
        checklistItem.parentNode.insertBefore(errorDetails, checklistItem.nextSibling);
    } else {
        console.error(`检查项 ${itemName} 的父节点不存在`);
    }
}

// 隐藏标题错误详情
function hideTitleErrorDetails(itemName) {
    const existingError = document.getElementById(`${itemName}-error-details`);
    if (existingError) {
        existingError.remove();
    }
}

// 只获取 docx 文档中 3.1 功能需求下的功能编号
function extractFunctionIdsFrom31(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const allLis = Array.from(doc.querySelectorAll('li'));
    let in31Section = false;
    let functionIds = [];

    for (let i = 0; i < allLis.length; i++) {
        const text = allLis[i].textContent.trim();
        // 判断是否进入3.1功能需求段落
        if (/^3\.1[\s．、.]*功能需求/.test(text)) {
            in31Section = true;
            continue;
        }
        // 判断是否离开3.1功能需求段落（遇到3.2或3.等）
        if (in31Section && /^3\.[2-9]/.test(text)) {
            break;
        }
        // 在3.1功能需求段落内，提取功能编号
        if (in31Section && /功能编号[：:]/.test(text)) {
            // 例如：功能编号：ABC123
            const match = text.match(/功能编号[：:](.+)/);
            if (match && match[1]) {
                functionIds.push(match[1].trim());
            }
        }
    }
    return functionIds;
}

// 功能编号和名称是否匹配 检查逻辑
function checkFuncMatch() {
    const checkBtn = document.getElementById('checkFuncMatchBtn');
    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];
    if (!requirementFile || !designFile) {
        alert('请先上传需求文档和概要设计文档');
        return;
    }
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';
    readFile(requirementFile)
        .then(requirementHtml => {
            return readFile(designFile)
                .then(designHtml => {
                    const requirementInfo = extractInfo(requirementHtml, 'li');
                    const designInfo = extractInfo(designHtml, 'td');
                    const errors = compareInfo(requirementInfo, designInfo);
                    // 自动标记检查项
                    const isPass = errors.length === 0;
                    updateChecklistStatus('funcMatch', isPass);
                    // 只显示到 result 区域，不再同步缓存
                    displayResult(errors, requirementInfo, designInfo);
                    checkBtn.disabled = false;
                    checkBtn.textContent = '检查';
                });
        })
        .catch(error => {
            alert('比对失败，请检查文档格式');
            checkBtn.disabled = false;
            checkBtn.textContent = '检查';
        });
}

// 功能模块检查逻辑（自定义比对）
function checkModuleMatch() {
    const checkBtn = document.getElementById('checkModuleBtn');
    const designFile = document.getElementById('designFile').files[0];
    if (!designFile) {
        alert('请先上传概要设计文档');
        return;
    }
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';
    readFile(designFile)
        .then(designHtml => {
            // 1. 读取第4到倒数第2个<table>，每个表格只提取第1行第4列为模块编号key，第2列为模块名称value
            const parser = new DOMParser();
            const doc = parser.parseFromString(designHtml, 'text/html');
            const tables = doc.querySelectorAll('table');
            let moduleMapA = {};
            function extractModuleIdAndNameFromDetailTable(table) {
                const rows = table.querySelectorAll('tr');
                if (rows.length > 0) {
                    const tds = rows[0].querySelectorAll('td');
                    if (tds.length >= 4) {
                        const moduleName = tds[1].textContent.trim();
                        const moduleId = tds[3].textContent.trim();
                        if (moduleId && moduleName && moduleId !== '模块编号') {
                            return { moduleId, moduleName };
                        }
                    }
                }
                return null;
            }
            if (tables.length >= 5) {
                for (let i = 3; i < tables.length - 2; i++) { // 0-based: 第4到倒数第2个
                    const detail = extractModuleIdAndNameFromDetailTable(tables[i]);
                    if (detail) {
                        moduleMapA[detail.moduleId] = detail.moduleName;
                    }
                }
            }
            // 2. 读取最后一个<table>，第2列为编号，第3列为名称，兼容rowspan
            let moduleMapB = {};
            if (tables.length > 0) {
                const lastTable = tables[tables.length - 1];
                const rows = lastTable.querySelectorAll('tr');
                let lastModuleId = '';
                let lastModuleName = '';
                rows.forEach((row, idx) => {
                    const tds = row.querySelectorAll('td');
                    if (tds.length >= 4) {
                        // 跳过表头
                        if (idx === 0) return;
                        let moduleId = tds[2].textContent.trim();
                        let moduleName = tds[3].textContent.trim();
                        if (!moduleId) moduleId = lastModuleId;
                        if (!moduleName) moduleName = lastModuleName;
                        if (moduleId === '模块编号') return;
                        if (moduleId && moduleName) {
                            moduleMapB[moduleId] = moduleName;
                        }
                        if (moduleId) lastModuleId = moduleId;
                        if (moduleName) lastModuleName = moduleName;
                    }
                });
            }
            // 3. 比对
            let errors = [];
            for (const [id, name] of Object.entries(moduleMapA)) {
                if (!(id in moduleMapB)) {
                    errors.push(`模块编号 ${id} 在汇总表中未找到。`);
                } else if (moduleMapB[id] !== name) {
                    errors.push(`模块编号 ${id} 名称不一致：表格为 ${name}，汇总表为 ${moduleMapB[id]}。`);
                }
            }
            for (const id in moduleMapB) {
                if (!(id in moduleMapA)) {
                    errors.push(`模块编号 ${id} 在模块表中未找到。`);
                }
            }
            // 自动标记检查项
            const isPass = errors.length === 0;
            updateChecklistStatus('moduleCheck', isPass);
            // 显示详细结果
            showModuleCheckResult(errors, moduleMapA, moduleMapB);
            checkBtn.disabled = false;
            checkBtn.textContent = '检查';
        })
        .catch(error => {
            alert('比对失败，请检查文档格式');
            checkBtn.disabled = false;
            checkBtn.textContent = '检查';
        });
}

// 显示功能模块比对结果
function showModuleCheckResult(errors, moduleMapA, moduleMapB) {
    let html = '';
    // 先输出比对结果
    if (errors.length === 0) {
        html += "<div class='error-details'><div class='error-title' style='color:green;'>模块比对成功！</div></div>";
    } else {
        html += `<div class='error-details'><div class='error-title'>模块比对结果：</div><ul style='margin:0;padding-left:18px;'>${errors.map(e => `<li>${e}</li>`).join("")}</ul></div>`;
    }
    // 再输出功能模块表
    html += '<div class="error-details">';
    html += '<div class="error-title">功能模块表：</div>';
    html += '<div class="error-item">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr><th style="text-align:left;padding:2px 6px;">模块编号</th><th style="text-align:left;padding:2px 6px;">模块名称</th></tr>';
    for (const [id, name] of Object.entries(moduleMapA)) {
        html += `<tr><td style="padding:2px 6px;border-bottom:1px solid #eee;">${id}</td><td style="padding:2px 6px;border-bottom:1px solid #eee;">${name}</td></tr>`;
    }
    html += '</table>';
    html += '</div>';
    html += '</div>';
    // 再输出功能模块汇总表
    html += '<div class="error-details">';
    html += '<div class="error-title">功能模块汇总表：</div>';
    html += '<div class="error-item">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr><th style="text-align:left;padding:2px 6px;">模块编号</th><th style="text-align:left;padding:2px 6px;">模块名称</th></tr>';
    for (const [id, name] of Object.entries(moduleMapB)) {
        html += `<tr><td style="padding:2px 6px;border-bottom:1px solid #eee;">${id}</td><td style="padding:2px 6px;border-bottom:1px solid #eee;">${name}</td></tr>`;
    }
    html += '</table>';
    html += '</div>';
    html += '</div>';
    // 在功能模块检查按钮下方显示
    const btn = document.getElementById('checkModuleBtn');
    let resultDiv = document.getElementById('moduleResult');
    if (!resultDiv) {
        resultDiv = document.createElement('div');
        resultDiv.id = 'moduleResult';
        btn.parentNode.insertBefore(resultDiv, btn.nextSibling);
    }
    resultDiv.innerHTML = html;
}

// 固定人员名单
const fixedPersons = [
    "阎辉", "周锐", "李林珍", "见浩亮", "万铸发", "武常钦", "张云朋", "武国军",
    "张丽娜", "李宁", "刘葛亮", "郭文杰", "张秋艳", "宋添天"
];

function checkPersonMatch() {
    const checkBtn = document.getElementById('checkPersonBtn');
    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];
    if (!requirementFile || !designFile) {
        alert('请先上传需求文档和概要设计文档');
        return;
    }
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';

    Promise.all([readFile(requirementFile), readFile(designFile)]).then(([reqHtml, desHtml]) => {
        // 需求文档
        const reqDoc = new DOMParser().parseFromString(reqHtml, 'text/html');
        const reqTables = reqDoc.querySelectorAll('table');
        let reqDraft = [], reqRevise = [], reqReview = [];
        if (reqTables.length >= 2) {
            // table 1, 第一行第三列
            const t1r1 = reqTables[0].querySelectorAll('tr')[0];
            if (t1r1) {
                const tds = t1r1.querySelectorAll('td');
                if (tds.length >= 3) {
                    const match = tds[2].textContent.match(/起草人[:：]?(.*?)(起草日期|$)/);
                    if (match) reqDraft = match[1].replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                }
            }
            // table 2, 第二行第4/5列
            const t2r2 = reqTables[1].querySelectorAll('tr')[2];
            if (t2r2) {
                const tds = t2r2.querySelectorAll('td');
                if (tds.length >= 5) {
                    reqRevise = tds[3].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                    reqReview = tds[4].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                }
            }
        }
        // 概要设计文档
        const desDoc = new DOMParser().parseFromString(desHtml, 'text/html');
        const desTables = desDoc.querySelectorAll('table');
        let desDraft = [], desRevise = [], desReview = [];
        if (desTables.length >= 2) {
            // table 1, 第一行第三列
            const t1r1 = desTables[0].querySelectorAll('tr')[0];
            if (t1r1) {
                const tds = t1r1.querySelectorAll('td');
                if (tds.length >= 3) {
                    const match = tds[2].textContent.match(/起草人[:：]?(.*?)(起草日期|$)/);
                    if (match) desDraft = match[1].replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                }
            }
            // table 2, 第二行第4/5列
            const t2r2 = desTables[1].querySelectorAll('tr')[2];
            if (t2r2) {
                const tds = t2r2.querySelectorAll('td');
                if (tds.length >= 5) {
                    desRevise = tds[3].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                    desReview = tds[4].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                }
            }
        }
        // 检查
        function checkPersons(arr) {
            return arr && arr.length ? arr.every(name => fixedPersons.includes(name)) : false;
        }
        const reqDraftPass = checkPersons(reqDraft);
        const reqRevisePass = checkPersons(reqRevise);
        const reqReviewPass = checkPersons(reqReview);
        const desDraftPass = checkPersons(desDraft);
        const desRevisePass = checkPersons(desRevise);
        const desReviewPass = checkPersons(desReview);

        const isPass = reqDraftPass && reqRevisePass && reqReviewPass && desDraftPass && desRevisePass && desReviewPass;
        updateChecklistStatus('fillerValid', isPass);

        // 展示
        let html = '';
        html += `<div class="error-details"><div class="error-title">需求文档：</div>
        <div class="error-item">起草人：${reqDraft ? reqDraft.join('，') : ''} ${reqDraftPass ? '✓' : '✗'}</div>
        <div class="error-item">修订人：${reqRevise ? reqRevise.join('，') : ''} ${reqRevisePass ? '✓' : '✗'}</div>
        <div class="error-item">复核人：${reqReview ? reqReview.join('，') : ''} ${reqReviewPass ? '✓' : '✗'}</div>
        </div>`;
        html += `<div class="error-details"><div class="error-title">概要设计文档：</div>
        <div class="error-item">起草人：${desDraft ? desDraft.join('，') : ''} ${desDraftPass ? '✓' : '✗'}</div>
        <div class="error-item">修订人：${desRevise ? desRevise.join('，') : ''} ${desRevisePass ? '✓' : '✗'}</div>
        <div class="error-item">复核人：${desReview ? desReview.join('，') : ''} ${desReviewPass ? '✓' : '✗'}</div>
        </div>`;
        html += `<div class="error-details"><div class="error-title">固定人员：</div>
        <div class="error-item">${fixedPersons.join('，')}</div></div>`;
        if (!isPass) {
            html = `<div class="error-details"><div class="error-title" style="color:#dc3545;">人员检查未通过！</div></div>` + html;
        } else {
            html = `<div class="error-details"><div class="error-title" style="color:green;">人员检查通过！</div></div>` + html;
        }
        let resultDiv = document.getElementById('personResult');
        if (!resultDiv) {
            resultDiv = document.createElement('div');
            resultDiv.id = 'personResult';
            checkBtn.parentNode.insertBefore(resultDiv, checkBtn.nextSibling);
        }
        resultDiv.innerHTML = html;
        checkBtn.disabled = false;
        checkBtn.textContent = '检查';
    }).catch(() => {
        alert('检查失败，请检查文档格式');
        checkBtn.disabled = false;
        checkBtn.textContent = '检查';
    });
}

// 第6项：接口需求检查
function checkInterfaceRequirement() {
    const checkBtn = document.getElementById('checkInterfaceBtn');
    const requirementFile = document.getElementById('requirementFile').files[0];
    if (!requirementFile) {
        alert('请先上传需求文档');
        return;
    }
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';
    readFile(requirementFile).then(requirementHtml => {
        const parser = new DOMParser();
        const doc = parser.parseFromString(requirementHtml, 'text/html');
        // 查找<h2>外部接口需求</h2>下一个<p>
        let found = false;
        let resultText = '';
        const h2s = Array.from(doc.querySelectorAll('h2'));
        for (let h2 of h2s) {
            if (h2.textContent.trim().includes('外部接口需求')) {
                // 找到下一个兄弟节点且为<p>
                let next = h2.nextElementSibling;
                while (next && next.tagName !== 'P') next = next.nextElementSibling;
                if (next && next.tagName === 'P') {
                    resultText = next.textContent.trim();
                    found = true;
                }
                break;
            }
        }
        let notPassPattern = /^(无|不涉及|本需求不涉及外部接口)$/;
        let isPass = found && !notPassPattern.test(resultText);
        // 自动标记
        updateChecklistStatus('interfaceNone', isPass);
        // 展示
        let html = '';
        if (!found) {
            html = '<div class="error-details"><div class="error-title">未找到"外部接口需求"章节或其下的内容！</div></div>';
        } else if (!isPass) {
            html = `<div class="error-details"><div class="error-title" style="color:#dc3545;">接口需求检查未通过！</div><div class="error-item">内容：${resultText}</div></div>`;
        } else {
            html = `<div class="error-details"><div class="error-title" style="color:green;">接口需求检查通过！</div><div class="error-item">内容：${resultText}</div></div>`;
        }
        let resultDiv = document.getElementById('interfaceResult');
        if (!resultDiv) {
            resultDiv = document.createElement('div');
            resultDiv.id = 'interfaceResult';
            checkBtn.parentNode.insertBefore(resultDiv, checkBtn.nextSibling);
        }
        resultDiv.innerHTML = html;
        checkBtn.disabled = false;
        checkBtn.textContent = '检查';
    }).catch(() => {
        alert('检查失败，请检查文档格式');
        checkBtn.disabled = false;
        checkBtn.textContent = '检查';
    });
}

// 第7项：概要设计文档接口检查
function checkDesignInterfaceRequirement() {
    const checkBtn = document.getElementById('checkDesignInterfaceBtn');
    const designFile = document.getElementById('designFile').files[0];
    if (!designFile) {
        alert('请先上传概要设计文档');
        return;
    }
    checkBtn.disabled = true;
    checkBtn.textContent = '检查中...';
    readFile(designFile).then(designHtml => {
        const parser = new DOMParser();
        const doc = parser.parseFromString(designHtml, 'text/html');

        // 检查三个接口标题：用户接口、外部接口、内部接口
        const interfaceTypes = ['用户接口', '外部接口', '内部接口'];
        let results = [];
        let allPass = true;

        interfaceTypes.forEach(interfaceType => {
            let found = false;
            let resultText = '';
            const h2s = Array.from(doc.querySelectorAll('h2'));

            for (let h2 of h2s) {
                if (h2.textContent.trim().includes(interfaceType)) {
                    // 找到下一个兄弟节点且为<p>
                    let next = h2.nextElementSibling;
                    while (next && next.tagName !== 'P') next = next.nextElementSibling;
                    if (next && next.tagName === 'P') {
                        resultText = next.textContent.trim();
                        found = true;
                    }
                    break;
                }
            }

            let notPassPattern = /^(无|不涉及|本需求不涉及.*接口)$/;
            let isPass = found && !notPassPattern.test(resultText);

            results.push({
                type: interfaceType,
                found: found,
                text: resultText,
                isPass: isPass
            });

            if (!isPass) {
                allPass = false;
            }
        });

        // 自动标记
        updateChecklistStatus('designInterfaceNone', allPass);

        // 展示结果
        let html = '';
        if (!allPass) {
            html = `<div class="error-details"><div class="error-title" style="color:#dc3545;">概要设计接口检查未通过！</div></div>`;
        } else {
            html = `<div class="error-details"><div class="error-title" style="color:green;">概要设计接口检查通过！</div></div>`;
        }

        // 显示每个接口的检查结果
        results.forEach(result => {
            html += `<div class="error-details"><div class="error-title">${result.type}：</div>`;
            if (!result.found) {
                html += `<div class="error-item">未找到"${result.type}"章节或其下的内容</div>`;
            } else {
                html += `<div class="error-item">内容：${result.text} ${result.isPass ? '✓' : '✗'}</div>`;
            }
            html += '</div>';
        });

        let resultDiv = document.getElementById('designInterfaceResult');
        if (!resultDiv) {
            resultDiv = document.createElement('div');
            resultDiv.id = 'designInterfaceResult';
            checkBtn.parentNode.insertBefore(resultDiv, checkBtn.nextSibling);
        }
        resultDiv.innerHTML = html;
        checkBtn.disabled = false;
        checkBtn.textContent = '检查';
    }).catch(() => {
        alert('检查失败，请检查文档格式');
        checkBtn.disabled = false;
        checkBtn.textContent = '检查';
    });
}

// 一键检查功能
async function performOneClickCheck() {
    const oneClickBtn = document.getElementById('oneClickCheckBtn');
    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];
    const crNameInput = document.getElementById('crNameInput');
    const requirementTitleInput = document.getElementById('requirementTitleInput');

    // 检查文件是否上传
    if (!requirementFile || !designFile) {
        alert('请先上传需求文档和概要设计文档');
        return;
    }

    // 检查CR编号是否输入
    if (!crNameInput.value.trim()) {
        alert('请先输入CR编号');
        return;
    }

    // 检查需求标题是否输入
    if (!requirementTitleInput.value.trim()) {
        alert('请先输入需求标题');
        return;
    }

    // 禁用按钮并显示检查中状态
    oneClickBtn.disabled = true;
    oneClickBtn.classList.add('checking');
    oneClickBtn.textContent = '检查中...';

    try {
        // 1. 检查文档名称
        await checkCRNameAsync();
        await new Promise(resolve => setTimeout(resolve, 100));

        // 2. 检查需求标题
        await checkRequirementTitleAsync();
        await new Promise(resolve => setTimeout(resolve, 100));

        // 3. 检查功能编号和名称匹配
        await checkFuncMatchAsync();
        await new Promise(resolve => setTimeout(resolve, 100));

        // 4. 检查功能模块匹配
        await checkModuleMatchAsync();
        await new Promise(resolve => setTimeout(resolve, 100));

        // 5. 检查人员
        await checkPersonMatchAsync();
        await new Promise(resolve => setTimeout(resolve, 100));

        // 6. 检查需求文档接口
        await checkInterfaceRequirementAsync();
        await new Promise(resolve => setTimeout(resolve, 100));

        // 7. 检查概要设计文档接口
        await checkDesignInterfaceRequirementAsync();

        // 显示完成消息
        setTimeout(() => {
            oneClickBtn.textContent = '检查完成';
            oneClickBtn.classList.remove('checking');
            oneClickBtn.classList.add('success');

            // 2秒后恢复按钮状态
            setTimeout(() => {
                oneClickBtn.disabled = false;
                oneClickBtn.textContent = '一键检查';
                oneClickBtn.classList.remove('success');
            }, 2000);
        }, 500);

    } catch (error) {
        console.error('一键检查过程中出错:', error);
        alert('检查过程中出现错误，请重试');

        // 恢复按钮状态
        oneClickBtn.disabled = false;
        oneClickBtn.textContent = '一键检查';
        oneClickBtn.classList.remove('checking');
    }
}

// 异步版本的CR名称检查
async function checkCRNameAsync() {
    return new Promise((resolve) => {
        const crNameInput = document.getElementById('crNameInput');
        const requirementFile = document.getElementById('requirementFile').files[0];
        const designFile = document.getElementById('designFile').files[0];

        const crNumber = crNameInput.value.trim();
        const expectedRequirementName = `${crNumber}_远程银行系统_需求规格说明书.docx`;
        const expectedDesignName = `${crNumber}_远程银行系统_概要设计说明书.docx`;

        const requirementMatch = requirementFile.name === expectedRequirementName;
        const designMatch = designFile.name === expectedDesignName;
        const isAllMatch = requirementMatch && designMatch;

        console.log('CR名称检查结果:', { requirementMatch, designMatch, isAllMatch });

        updateChecklistStatus('docName', isAllMatch);

        // 总是显示结果，无论通过还是不通过
        showErrorDetails('docName', requirementMatch, designMatch, expectedRequirementName, expectedDesignName, requirementFile.name, designFile.name);

        resolve();
    });
}

// 异步版本的需求标题检查
async function checkRequirementTitleAsync() {
    return new Promise((resolve, reject) => {
        const titleInput = document.getElementById('requirementTitleInput');
        const requirementFile = document.getElementById('requirementFile').files[0];

        readFile(requirementFile)
            .then(requirementHtml => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(requirementHtml, 'text/html');
                let docTitle = '';
                const pList = doc.querySelectorAll('p');
                if (pList.length >= 2) {
                    docTitle = pList[1].textContent.trim();
                }
                const inputTitle = titleInput.value.trim() + '_远程银行系统';
                const isMatch = docTitle === inputTitle;

                updateChecklistStatus('requirementTitle', isMatch);

                // 总是显示结果，无论通过还是不通过
                showTitleErrorDetails('requirementTitle', docTitle, inputTitle);

                resolve();
            })
            .catch(error => {
                console.error('检查需求标题时出错:', error);
                reject(error);
            });
    });
}

// 异步版本的功能编号和名称匹配检查
async function checkFuncMatchAsync() {
    return new Promise((resolve, reject) => {
        const requirementFile = document.getElementById('requirementFile').files[0];
        const designFile = document.getElementById('designFile').files[0];

        readFile(requirementFile)
            .then(requirementHtml => {
                return readFile(designFile)
                    .then(designHtml => {
                        const requirementInfo = extractInfo(requirementHtml, 'li');
                        const designInfo = extractInfo(designHtml, 'td');
                        const errors = compareInfo(requirementInfo, designInfo);
                        const isPass = errors.length === 0;

                        updateChecklistStatus('funcMatch', isPass);
                        displayResult(errors, requirementInfo, designInfo);

                        resolve();
                    });
            })
            .catch(error => {
                console.error('检查功能匹配时出错:', error);
                reject(error);
            });
    });
}

// 异步版本的功能模块检查
async function checkModuleMatchAsync() {
    return new Promise((resolve, reject) => {
        const designFile = document.getElementById('designFile').files[0];

        readFile(designFile)
            .then(designHtml => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(designHtml, 'text/html');
                const tables = doc.querySelectorAll('table');
                let moduleMapA = {};

                function extractModuleIdAndNameFromDetailTable(table) {
                    const rows = table.querySelectorAll('tr');
                    if (rows.length > 0) {
                        const tds = rows[0].querySelectorAll('td');
                        if (tds.length >= 4) {
                            const moduleName = tds[1].textContent.trim();
                            const moduleId = tds[3].textContent.trim();
                            if (moduleId && moduleName && moduleId !== '模块编号') {
                                return { moduleId, moduleName };
                            }
                        }
                    }
                    return null;
                }

                if (tables.length >= 5) {
                    for (let i = 3; i < tables.length - 2; i++) {
                        const detail = extractModuleIdAndNameFromDetailTable(tables[i]);
                        if (detail) {
                            moduleMapA[detail.moduleId] = detail.moduleName;
                        }
                    }
                }

                let moduleMapB = {};
                if (tables.length > 0) {
                    const lastTable = tables[tables.length - 1];
                    const rows = lastTable.querySelectorAll('tr');
                    let lastModuleId = '';
                    let lastModuleName = '';
                    rows.forEach((row, idx) => {
                        const tds = row.querySelectorAll('td');
                        if (tds.length >= 4) {
                            if (idx === 0) return;
                            let moduleId = tds[2].textContent.trim();
                            let moduleName = tds[3].textContent.trim();
                            if (!moduleId) moduleId = lastModuleId;
                            if (!moduleName) moduleName = lastModuleName;
                            if (moduleId === '模块编号') return;
                            if (moduleId && moduleName) {
                                moduleMapB[moduleId] = moduleName;
                            }
                            if (moduleId) lastModuleId = moduleId;
                            if (moduleName) lastModuleName = moduleName;
                        }
                    });
                }

                let errors = [];
                for (const [id, name] of Object.entries(moduleMapA)) {
                    if (!(id in moduleMapB)) {
                        errors.push(`模块编号 ${id} 在汇总表中未找到。`);
                    } else if (moduleMapB[id] !== name) {
                        errors.push(`模块编号 ${id} 名称不一致：表格为 ${name}，汇总表为 ${moduleMapB[id]}。`);
                    }
                }
                for (const id in moduleMapB) {
                    if (!(id in moduleMapA)) {
                        errors.push(`模块编号 ${id} 在模块表中未找到。`);
                    }
                }

                const isPass = errors.length === 0;
                updateChecklistStatus('moduleCheck', isPass);
                showModuleCheckResult(errors, moduleMapA, moduleMapB);

                resolve();
            })
            .catch(error => {
                console.error('检查功能模块时出错:', error);
                reject(error);
            });
    });
}

// 异步版本的人员检查
async function checkPersonMatchAsync() {
    return new Promise((resolve, reject) => {
        const requirementFile = document.getElementById('requirementFile').files[0];
        const designFile = document.getElementById('designFile').files[0];

        Promise.all([readFile(requirementFile), readFile(designFile)])
            .then(([reqHtml, desHtml]) => {
                const reqDoc = new DOMParser().parseFromString(reqHtml, 'text/html');
                const reqTables = reqDoc.querySelectorAll('table');
                let reqDraft = [], reqRevise = [], reqReview = [];

                if (reqTables.length >= 2) {
                    const t1r1 = reqTables[0].querySelectorAll('tr')[0];
                    if (t1r1) {
                        const tds = t1r1.querySelectorAll('td');
                        if (tds.length >= 3) {
                            const match = tds[2].textContent.match(/起草人[:：]?(.*?)(起草日期|$)/);
                            if (match) reqDraft = match[1].replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                        }
                    }
                    const t2r2 = reqTables[1].querySelectorAll('tr')[2];
                    if (t2r2) {
                        const tds = t2r2.querySelectorAll('td');
                        if (tds.length >= 5) {
                            reqRevise = tds[3].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                            reqReview = tds[4].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                        }
                    }
                }

                const desDoc = new DOMParser().parseFromString(desHtml, 'text/html');
                const desTables = desDoc.querySelectorAll('table');
                let desDraft = [], desRevise = [], desReview = [];

                if (desTables.length >= 2) {
                    const t1r1 = desTables[0].querySelectorAll('tr')[0];
                    if (t1r1) {
                        const tds = t1r1.querySelectorAll('td');
                        if (tds.length >= 3) {
                            const match = tds[2].textContent.match(/起草人[:：]?(.*?)(起草日期|$)/);
                            if (match) desDraft = match[1].replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                        }
                    }
                    const t2r2 = desTables[1].querySelectorAll('tr')[2];
                    if (t2r2) {
                        const tds = t2r2.querySelectorAll('td');
                        if (tds.length >= 5) {
                            desRevise = tds[3].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                            desReview = tds[4].textContent.trim().replace(/[，,、]/g, ',').split(',').map(s => s.trim()).filter(Boolean);
                        }
                    }
                }

                function checkPersons(arr) {
                    return arr && arr.length ? arr.every(name => fixedPersons.includes(name)) : false;
                }

                const reqDraftPass = checkPersons(reqDraft);
                const reqRevisePass = checkPersons(reqRevise);
                const reqReviewPass = checkPersons(reqReview);
                const desDraftPass = checkPersons(desDraft);
                const desRevisePass = checkPersons(desRevise);
                const desReviewPass = checkPersons(desReview);

                const isPass = reqDraftPass && reqRevisePass && reqReviewPass && desDraftPass && desRevisePass && desReviewPass;
                updateChecklistStatus('fillerValid', isPass);

                let html = '';
                html += `<div class="error-details"><div class="error-title">需求文档：</div>
                <div class="error-item">起草人：${reqDraft ? reqDraft.join('，') : ''} ${reqDraftPass ? '✓' : '✗'}</div>
                <div class="error-item">修订人：${reqRevise ? reqRevise.join('，') : ''} ${reqRevisePass ? '✓' : '✗'}</div>
                <div class="error-item">复核人：${reqReview ? reqReview.join('，') : ''} ${reqReviewPass ? '✓' : '✗'}</div>
                </div>`;
                html += `<div class="error-details"><div class="error-title">概要设计文档：</div>
                <div class="error-item">起草人：${desDraft ? desDraft.join('，') : ''} ${desDraftPass ? '✓' : '✗'}</div>
                <div class="error-item">修订人：${desRevise ? desRevise.join('，') : ''} ${desRevisePass ? '✓' : '✗'}</div>
                <div class="error-item">复核人：${desReview ? desReview.join('，') : ''} ${desReviewPass ? '✓' : '✗'}</div>
                </div>`;
                html += `<div class="error-details"><div class="error-title">固定人员：</div>
                <div class="error-item">${fixedPersons.join('，')}</div></div>`;
                if (!isPass) {
                    html = `<div class="error-details"><div class="error-title" style="color:#dc3545;">人员检查未通过！</div></div>` + html;
                } else {
                    html = `<div class="error-details"><div class="error-title" style="color:green;">人员检查通过！</div></div>` + html;
                }

                let resultDiv = document.getElementById('personResult');
                if (!resultDiv) {
                    resultDiv = document.createElement('div');
                    resultDiv.id = 'personResult';
                    const checkBtn = document.getElementById('checkPersonBtn');
                    checkBtn.parentNode.insertBefore(resultDiv, checkBtn.nextSibling);
                }
                resultDiv.innerHTML = html;

                resolve();
            })
            .catch(error => {
                console.error('检查人员时出错:', error);
                reject(error);
            });
    });
}

// 异步版本的接口需求检查
async function checkInterfaceRequirementAsync() {
    return new Promise((resolve, reject) => {
        const requirementFile = document.getElementById('requirementFile').files[0];

        readFile(requirementFile)
            .then(requirementHtml => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(requirementHtml, 'text/html');
                let found = false;
                let resultText = '';
                const h2s = Array.from(doc.querySelectorAll('h2'));

                for (let h2 of h2s) {
                    if (h2.textContent.trim().includes('外部接口需求')) {
                        let next = h2.nextElementSibling;
                        while (next && next.tagName !== 'P') next = next.nextElementSibling;
                        if (next && next.tagName === 'P') {
                            resultText = next.textContent.trim();
                            found = true;
                        }
                        break;
                    }
                }

                let notPassPattern = /^(无|不涉及|本需求不涉及外部接口*)$/;
                let isPass = found && !notPassPattern.test(resultText);

                updateChecklistStatus('interfaceNone', isPass);

                let html = '';
                if (!found) {
                    html = '<div class="error-details"><div class="error-title">未找到"外部接口需求"章节或其下的内容！</div></div>';
                } else if (!isPass) {
                    html = `<div class="error-details"><div class="error-title" style="color:#dc3545;">接口需求检查未通过！</div><div class="error-item">内容：${resultText}</div></div>`;
                } else {
                    html = `<div class="error-details"><div class="error-title" style="color:green;">接口需求检查通过！</div><div class="error-item">内容：${resultText}</div></div>`;
                }

                let resultDiv = document.getElementById('interfaceResult');
                if (!resultDiv) {
                    resultDiv = document.createElement('div');
                    resultDiv.id = 'interfaceResult';
                    const checkBtn = document.getElementById('checkInterfaceBtn');
                    checkBtn.parentNode.insertBefore(resultDiv, checkBtn.nextSibling);
                }
                resultDiv.innerHTML = html;

                resolve();
            })
            .catch(error => {
                console.error('检查接口需求时出错:', error);
                reject(error);
            });
    });
}

// 异步版本的概要设计文档接口检查
async function checkDesignInterfaceRequirementAsync() {
    return new Promise((resolve, reject) => {
        const designFile = document.getElementById('designFile').files[0];

        readFile(designFile)
            .then(designHtml => {
                const parser = new DOMParser();
                const doc = parser.parseFromString(designHtml, 'text/html');

                const interfaceTypes = ['用户接口', '外部接口', '内部接口'];
                let results = [];
                let allPass = true;

                interfaceTypes.forEach(interfaceType => {
                    let found = false;
                    let resultText = '';
                    const h2s = Array.from(doc.querySelectorAll('h2'));

                    for (let h2 of h2s) {
                        if (h2.textContent.trim().includes(interfaceType)) {
                            let next = h2.nextElementSibling;
                            while (next && next.tagName !== 'P') next = next.nextElementSibling;
                            if (next && next.tagName === 'P') {
                                resultText = next.textContent.trim();
                                found = true;
                            }
                            break;
                        }
                    }

                    let notPassPattern = /^(无|不涉及|本需求不涉及.*接口*)$/;
                    let isPass = found && !notPassPattern.test(resultText);

                    results.push({
                        type: interfaceType,
                        found: found,
                        text: resultText,
                        isPass: isPass
                    });

                    if (!isPass) {
                        allPass = false;
                    }
                });

                updateChecklistStatus('designInterfaceNone', allPass);

                let html = '';
                if (!allPass) {
                    html = `<div class="error-details"><div class="error-title" style="color:#dc3545;">概要设计接口检查未通过！</div></div>`;
                } else {
                    html = `<div class="error-details"><div class="error-title" style="color:green;">概要设计接口检查通过！</div></div>`;
                }

                results.forEach(result => {
                    html += `<div class="error-details"><div class="error-title">${result.type}：</div>`;
                    if (!result.found) {
                        html += `<div class="error-item">未找到"${result.type}"章节或其下的内容</div>`;
                    } else {
                        html += `<div class="error-item">内容：${result.text} ${result.isPass ? '✓' : '✗'}</div>`;
                    }
                    html += '</div>';
                });

                let resultDiv = document.getElementById('designInterfaceResult');
                if (!resultDiv) {
                    resultDiv = document.createElement('div');
                    resultDiv.id = 'designInterfaceResult';
                    const checkBtn = document.getElementById('checkDesignInterfaceBtn');
                    checkBtn.parentNode.insertBefore(resultDiv, checkBtn.nextSibling);
                }
                resultDiv.innerHTML = html;

                resolve();
            })
            .catch(error => {
                console.error('检查概要设计接口时出错:', error);
                reject(error);
            });
    });
}

// 导出检查结果功能
function exportCheckResults(format) {
    const exportFormatSelect = document.getElementById('exportFormat');

    // 检查是否已经进行了检查
    const checklistItems = document.querySelectorAll('.checklist-item');
    let hasResults = false;

    checklistItems.forEach(item => {
        const statusIcon = item.querySelector('.status-icon');
        if (statusIcon) {
            hasResults = true;
        }
    });

    if (!hasResults) {
        // 使用更友好的提示方式
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: #dc3545;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(220, 53, 69, 0.3);
            z-index: 1000;
            font-size: 14px;
            animation: slideIn 0.3s ease-out;
        `;
        notification.textContent = '请先进行一键检查，然后再导出结果';
        document.body.appendChild(notification);

        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease-in';
            setTimeout(() => notification.remove(), 300);
        }, 3000);

        return;
    }

    try {
        // 收集检查结果
        const results = collectCheckResults();

        // 生成导出内容
        const exportContent = format === 'csv' ?
            generateCSVContent(results) :
            generateExportContent(results);

        // 创建并下载文件
        downloadFile(exportContent, format);

        // 显示成功提示
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: #28a745;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(40, 167, 69, 0.3);
            z-index: 1000;
            font-size: 14px;
            animation: slideIn 0.3s ease-out;
        `;
        notification.textContent = `检查结果已成功导出为 ${format.toUpperCase()} 格式！`;
        document.body.appendChild(notification);

        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease-in';
            setTimeout(() => notification.remove(), 300);
        }, 3000);

    } catch (error) {
        console.error('导出失败:', error);

        // 显示错误提示
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: #dc3545;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(220, 53, 69, 0.3);
            z-index: 1000;
            font-size: 14px;
            animation: slideIn 0.3s ease-out;
        `;
        notification.textContent = '导出失败，请重试';
        document.body.appendChild(notification);

        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease-in';
            setTimeout(() => notification.remove(), 300);
        }, 3000);
    }
}

// 收集检查结果
function collectCheckResults() {
    const results = [];
    const crNameInput = document.getElementById('crNameInput');
    const requirementTitleInput = document.getElementById('requirementTitleInput');

    // 检查项目配置
    const checkItems = [
        {
            name: 'docName',
            title: '1. 文档名称是否正确',
            description: '检查需求文档和概要设计文档的文件名是否符合规范'
        },
        {
            name: 'requirementTitle',
            title: '2. 需求标题是否正确',
            description: '检查需求文档中的标题是否与输入的标题一致'
        },
        {
            name: 'funcMatch',
            title: '3. 功能编号和名称是否匹配',
            description: '检查需求文档和概要设计文档中的功能编号和名称是否一致'
        },
        {
            name: 'moduleCheck',
            title: '4. 功能模块是否匹配',
            description: '检查概要设计文档中的功能模块表与汇总表是否一致'
        },
        {
            name: 'fillerValid',
            title: '5. 填写人员是否符合要求',
            description: '检查文档中的起草人、修订人、复核人是否在固定人员名单中'
        },
        {
            name: 'interfaceNone',
            title: '6. 需求文档中接口设计是否存在"无"或"不涉及"',
            description: '检查需求文档的外部接口需求章节是否包含"无"或"不涉及"'
        },
        {
            name: 'designInterfaceNone',
            title: '7. 概要设计文档中接口是否存在"无"或"不涉及"',
            description: '检查概要设计文档的用户接口、外部接口、内部接口章节是否包含"无"或"不涉及"'
        }
    ];

    // 收集每个检查项的结果
    checkItems.forEach(item => {
        const radioBtn = document.querySelector(`input[name="${item.name}"]:checked`);
        const statusIcon = document.querySelector(`input[name="${item.name}"]`).closest('.checklist-item').querySelector('.status-icon');
        const errorDetails = document.getElementById(`${item.name}-error-details`);

        let status = '未检查';
        let result = '';

        if (radioBtn) {
            status = radioBtn.value === 'yes' ? '通过' : '不通过';

            if (statusIcon) {
                result = statusIcon.textContent === '✓' ? '通过' : '不通过';
            }

            // 获取详细结果
            if (errorDetails) {
                result = errorDetails.textContent.trim();
            } else if (status === '通过') {
                // 根据检查项类型设置通过时的结果描述
                switch (item.name) {
                    case 'docName':
                        result = '文档名称符合规范';
                        break;
                    case 'requirementTitle':
                        result = '需求标题正确';
                        break;
                    case 'funcMatch':
                        result = '功能编号和名称匹配';
                        break;
                    case 'moduleCheck':
                        result = '功能模块匹配';
                        break;
                    case 'fillerValid':
                        result = '填写人员符合要求';
                        break;
                    case 'interfaceNone':
                        result = '接口需求描述正确';
                        break;
                    case 'designInterfaceNone':
                        result = '接口设计描述正确';
                        break;
                    default:
                        result = '检查通过';
                }
            }
        }

        results.push({
            title: item.title,
            description: item.description,
            status: status,
            result: result
        });
    });

    return results;
}

// 生成导出内容
function generateExportContent(results) {
    const crNameInput = document.getElementById('crNameInput');
    const requirementTitleInput = document.getElementById('requirementTitleInput');
    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];

    const now = new Date();
    const timestamp = now.toLocaleString('zh-CN');

    let content = `文档比对检查结果报告\n`;
    content += `生成时间：${timestamp}\n`;
    content += `CR编号：${crNameInput.value.trim() || '未输入'}\n`;
    content += `需求标题：${requirementTitleInput.value.trim() || '未输入'}\n`;
    content += `需求文档：${requirementFile ? requirementFile.name : '未上传'}\n`;
    content += `概要设计文档：${designFile ? designFile.name : '未上传'}\n`;
    content += `\n`;
    content += `检查项目\t是否通过\t检查结果\n`;
    content += `─`.repeat(80) + `\n`;

    results.forEach(item => {
        content += `${item.title}\t${item.status}\t${item.result}\n`;
    });

    content += `\n`;
    content += `检查项目详细说明：\n`;
    content += `─`.repeat(80) + `\n`;

    results.forEach(item => {
        content += `${item.title}\n`;
        content += `说明：${item.description}\n`;
        content += `状态：${item.status}\n`;
        content += `结果：${item.result}\n`;
        content += `\n`;
    });

    return content;
}

// 生成CSV格式内容
function generateCSVContent(results) {
    const crNameInput = document.getElementById('crNameInput');
    const requirementTitleInput = document.getElementById('requirementTitleInput');
    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];

    const now = new Date();
    const timestamp = now.toLocaleString('zh-CN');

    // CSV头部信息
    let csvContent = `文档比对检查结果报告\n`;
    csvContent += `生成时间,${timestamp}\n`;
    csvContent += `CR编号,${crNameInput.value.trim() || '未输入'}\n`;
    csvContent += `需求标题,${requirementTitleInput.value.trim() || '未输入'}\n`;
    csvContent += `需求文档,${requirementFile ? requirementFile.name : '未上传'}\n`;
    csvContent += `概要设计文档,${designFile ? designFile.name : '未上传'}\n`;
    csvContent += `\n`;

    // CSV表头
    csvContent += `检查项目,是否通过,检查结果,详细说明\n`;

    // CSV数据行
    results.forEach(item => {
        // 处理CSV中的特殊字符（逗号、引号、换行符）
        const title = escapeCSVField(item.title);
        const status = escapeCSVField(item.status);
        const result = escapeCSVField(item.result);
        const description = escapeCSVField(item.description);

        csvContent += `${title},${status},${result},${description}\n`;
    });

    return csvContent;
}

// 转义CSV字段中的特殊字符
function escapeCSVField(field) {
    if (!field) return '';

    // 如果字段包含逗号、引号或换行符，需要用引号包围
    if (field.includes(',') || field.includes('"') || field.includes('\n') || field.includes('\r')) {
        // 将字段中的引号替换为两个引号
        const escapedField = field.replace(/"/g, '""');
        return `"${escapedField}"`;
    }

    return field;
}

// 下载文件
function downloadFile(content, format) {
    const crNameInput = document.getElementById('crNameInput');
    const crNumber = crNameInput.value.trim() || '未知';
    const now = new Date();
    const dateStr = now.toISOString().split('T')[0];
    const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-');

    const filename = `检查结果_${crNumber}_${dateStr}_${timeStr}.${format}`;

    const blob = new Blob([content], { type: `text/${format};charset=utf-8` });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
