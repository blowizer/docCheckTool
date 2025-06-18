document.getElementById("requirementFile").addEventListener("change", function () {
    document.getElementById("requirementFileName").value = this.files[0]?.name || "";
});

document.getElementById("designFile").addEventListener("change", function () {
    document.getElementById("designFileName").value = this.files[0]?.name || "";
});

document.getElementById("compareBtn").addEventListener("click", function () {
    document.getElementById("result").innerHTML = ""; // 清空之前的比对结果

    const requirementFile = document.getElementById('requirementFile').files[0];
    const designFile = document.getElementById('designFile').files[0];

    if (!requirementFile || !designFile) {
        document.getElementById('result').innerHTML = '<div style="color:red;">请上传需求文档和概要设计文档。</div>';
        return;
    }
    if (!requirementFile.name.includes('需求') || !designFile.name.includes('概要')) {
        document.getElementById('result').innerHTML = '<div class="error">文件名不符合要求，请确保选择正确的文档。</div>';
        return;
    }


    Promise.all([readFile(requirementFile), readFile(designFile)])
        .then(([requirementHtml, designHtml]) => {
            const requirementInfo = extractInfo(requirementHtml, 'li');
            const designInfo = extractInfo(designHtml, 'td');
            const errors = compareInfo(requirementInfo, designInfo);
            displayResult(errors);
        })
        .catch(error => {
            console.error('读取文件时出错:', error);
            document.getElementById('result').innerHTML = '<div style="color:red;">读取文件时出错，请重试。</div>';
        });
});

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
            const arrayBuffer = e.target.result;
            mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
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
    const info = {};

    elements.forEach((element) => {
        const text = element.textContent.trim();
        if (selector === 'li' && text.includes('功能编号')) {
            const functionId = text.split(/[:：]/).pop().trim();
            const functionNameElement = element.nextElementSibling;
            if (functionNameElement) {
                const functionName = functionNameElement.textContent.split(/[:：]/).pop().trim();
                info[functionId] = functionName;
            }
        } else if (selector === 'td') {
            const rows = doc.querySelectorAll('tr');
            rows.forEach((row, rowIndex) => {
                if (rowIndex === 0) return;
                const cells = row.querySelectorAll('td');
                if (cells.length >= 2) {
                    const key = cells[0].textContent.trim();
                    const value = cells[1].textContent.trim();
                    if (key && value) {
                        info[key] = value;
                    }
                }
            });
        }
    });

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

function displayResult(errors) {
    document.getElementById("result").innerHTML = errors.length === 0
        ? "<div style='color:green;'>比对成功！</div>"
        : `<ul>${errors.map(e => `<li>${e}</li>`).join("")}</ul>`;
}
