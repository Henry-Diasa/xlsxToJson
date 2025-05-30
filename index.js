const fs = require('fs-extra');
const XLSX = require('xlsx');
const path = require('path');

function cleanTranslation(text) {
    if (!text) return '';
    return String(text).trim().replace(/\s+/g, ' ').replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
}

function abbreviate(word, len = 2) {
    return word ? word.slice(0, len) : 'x';
}

function generateKey(prefix, text, fallback, existingKeys) {
    let parts = [];
    if (text) {
        parts = text.toLowerCase().replace(/[^a-z0-9 ]/g, '').split(' ').filter(Boolean).slice(0, 2);
    }
    if (parts.length < 2 && fallback) {
        const fallbackParts = fallback.toLowerCase().replace(/[^a-z0-9 ]/g, '').split(' ').filter(Boolean);
        while (parts.length < 2 && fallbackParts.length > 0) {
            parts.push(fallbackParts.shift());
        }
    }
    while (parts.length < 2) {
        parts.push('xx');
    }
    // 简写为2个字母
    let part1 = abbreviate(parts[0], 2);
    let part2 = abbreviate(parts[1], 2);
    let key = `${prefix}.${part1}.${part2}`;
    // 长度控制
    if (key.length > 16) {
        part1 = abbreviate(parts[0], 1);
        part2 = abbreviate(parts[1], 1);
        key = `${prefix}.${part1}.${part2}`;
    }
    // 冲突处理
    let counter = 1;
    let finalKey = key;
    while (existingKeys.has(finalKey) || finalKey.length > 16) {
        finalKey = `${key}${counter}`;
        if (finalKey.length > 16) finalKey = finalKey.slice(0, 16);
        counter++;
    }
    existingKeys.add(finalKey);
    return finalKey;
}

function processXlsxFile(xlsxPath) {
    const fileBase = path.basename(xlsxPath, path.extname(xlsxPath));
    const prefix = fileBase.slice(0, 2).toLowerCase();
    const outputDir = path.join(process.cwd(), 'translations', fileBase);
    fs.ensureDirSync(outputDir);

    try {
        const workbook = XLSX.readFile(xlsxPath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (data.length < 2) {
            console.log(`${xlsxPath} 文件格式不正确：至少需要包含表头和一行数据`);
            return;
        }

        const headers = data[0];
        const langIndexes = {};
        headers.forEach((header, idx) => {
            langIndexes[header] = idx;
        });

        const enCol = Object.keys(langIndexes).find(h => h.toLowerCase().includes('en'));
        const zhCol = Object.keys(langIndexes).find(h => h.includes('简体'));

        const translations = {};
        Object.keys(langIndexes).forEach(lang => {
            translations[lang] = {};
        });

        const existingKeys = new Set();
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            const enText = enCol ? row[langIndexes[enCol]] : '';
            const zhText = zhCol ? row[langIndexes[zhCol]] : '';
            const key = generateKey(prefix, enText, zhText, existingKeys);
            Object.keys(langIndexes).forEach(lang => {
                const value = cleanTranslation(row[langIndexes[lang]]);
                translations[lang][key] = value || "";
            });
        }

        Object.values(translations).forEach(data => {
            existingKeys.forEach(key => {
                if (!(key in data)) {
                    data[key] = "";
                }
            });
        });

        Object.entries(translations).forEach(([lang, data]) => {
            // 只要有一个 value 非空就生成
            const hasContent = Object.values(data).some(v => v !== "");
            if (!hasContent) {
                console.log(`跳过生成 ${fileBase}/${lang}.json（内容全为空）`);
                return;
            }
            const outputPath = path.join(outputDir, `${lang}.json`);
            fs.writeJsonSync(outputPath, data, { spaces: 2 });
            console.log(`已生成 ${fileBase}/${lang}.json`);
        });
        console.log(`${xlsxPath} 转换完成！`);
    } catch (error) {
        console.error(`处理${xlsxPath}时出错:`, error.message);
    }
}

// 主入口：批量处理所有xlsx文件
if (process.argv.length === 2) {
    const files = fs.readdirSync('.').filter(f => f.endsWith('.xlsx'));
    if (files.length === 0) {
        console.log('当前目录下没有xlsx文件');
        process.exit(0);
    }
    files.forEach(file => processXlsxFile(file));
} else {
    // 兼容单文件处理
    processXlsxFile(process.argv[2]);
} 