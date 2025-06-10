const fs = require('fs-extra');
const XLSX = require('xlsx');
const path = require('path');

function cleanTranslation(text) {
    if (!text) return '';
    return String(text).trim().replace(/\s+/g, ' ').replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
}

// 新增：在单元格中查找匹配的JSON值并返回对应的key
function findMatchingKeys(cellText, baseJson) {
    if (!cellText) return [];
    
    const cleanCell = cleanTranslation(cellText);
    const matches = [];
    
    // 遍历所有JSON key-value对
    Object.entries(baseJson).forEach(([key, value]) => {
        const cleanValue = cleanTranslation(value);
        if (cleanValue && cleanValue.length > 2 && cleanCell.includes(cleanValue)) {
            matches.push({
                key: key,
                value: cleanValue,
                originalValue: value,
                startIndex: cleanCell.indexOf(cleanValue),
                endIndex: cleanCell.indexOf(cleanValue) + cleanValue.length
            });
        }
    });
    
    // 按长度排序，优先保留较长的匹配
    matches.sort((a, b) => b.value.length - a.value.length);
    
    // 过滤重叠的匹配
    const nonOverlappingMatches = [];
    for (const match of matches) {
        const hasOverlap = nonOverlappingMatches.some(existing => 
            (match.startIndex < existing.endIndex && match.endIndex > existing.startIndex)
        );
        if (!hasOverlap) {
            nonOverlappingMatches.push(match);
        }
    }
    
    return nonOverlappingMatches;
}

// 智能分割多值翻译文本
function splitMultipleTranslations(englishText, translationText, matches) {
    if (!translationText || matches.length <= 1) {
        return matches.length === 1 ? { [matches[0].key]: translationText } : {};
    }
    
    // 按英文文本中的位置排序
    const sortedMatches = matches.sort((a, b) => a.startIndex - b.startIndex);
    const result = {};
    
    // 尝试按比例和位置分割翻译文本
    const cleanEnglish = cleanTranslation(englishText);
    const cleanTranslationText = cleanTranslation(translationText);
    
    // 计算每个匹配项在英文中的相对位置和长度
    let lastIndex = 0;
    const segments = [];
    
    for (let i = 0; i < sortedMatches.length; i++) {
        const match = sortedMatches[i];
        const nextMatch = sortedMatches[i + 1];
        
        // 计算当前匹配项的相对起始位置
        const relativeStart = match.startIndex / cleanEnglish.length;
        const relativeLength = match.value.length / cleanEnglish.length;
        
        // 计算在翻译文本中的大概起始位置
        const translationStart = Math.floor(relativeStart * cleanTranslationText.length);
        
        let translationEnd;
        if (nextMatch) {
            // 如果有下一个匹配项，计算到下一个项开始前的位置
            const nextRelativeStart = nextMatch.startIndex / cleanEnglish.length;
            translationEnd = Math.floor(nextRelativeStart * cleanTranslationText.length);
        } else {
            // 如果是最后一个，延伸到文本末尾
            translationEnd = cleanTranslationText.length;
        }
        
        // 提取对应的翻译片段
        let segment = cleanTranslationText.substring(translationStart, translationEnd).trim();
        
        // 尝试智能分割：寻找可能的分隔符
        if (i < sortedMatches.length - 1) {
            // 寻找可能的句子边界或分隔符
            const boundaries = ['. ', '。', '; ', '；', ' | ', '｜'];
            for (const boundary of boundaries) {
                const boundaryIndex = segment.lastIndexOf(boundary);
                if (boundaryIndex > segment.length * 0.3) { // 边界不能太靠前
                    segment = segment.substring(0, boundaryIndex).trim();
                    break;
                }
            }
        }
        
        if (segment) {
            result[match.key] = segment;
            console.log(`    分割结果 "${match.key}": "${segment}"`);
        }
    }
    
    return result;
}

function processXlsxFile(xlsxPath) {
    const fileBase = path.basename(xlsxPath, path.extname(xlsxPath));
    const outputDir = path.join(process.cwd(), 'translations', fileBase);
    fs.ensureDirSync(outputDir);

    try {
        // 读取与Excel同名的json作为基准
        const jsonPath = path.join(process.cwd(), `${fileBase}.json`);
        const baseJson = fs.readJsonSync(jsonPath);

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
        if (!enCol) {
            console.log(`${xlsxPath} 缺少英文列`);
            return;
        }

        // 检查每种语言是否有内容
        const hasContent = {};
        Object.keys(langIndexes).forEach(lang => {
            hasContent[lang] = false;
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (row && row.length > langIndexes[lang] && cleanTranslation(row[langIndexes[lang]])) {
                    hasContent[lang] = true;
                    break;
                }
            }
        });

        // 创建翻译映射
        const keyTranslations = new Map(); // key -> {lang: translation}
        
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            
            const enValue = cleanTranslation(row[langIndexes[enCol]]);
            if (!enValue) continue;

            // 在当前行的英文内容中查找匹配的JSON key
            const matchingKeys = findMatchingKeys(enValue, baseJson);
            
            if (matchingKeys.length > 0) {
                console.log(`在行 ${i + 1} 中找到 ${matchingKeys.length} 个匹配的key:`);
                
                if (matchingKeys.length === 1) {
                    // 单个匹配，直接处理
                    const match = matchingKeys[0];
                    console.log(`  - "${match.key}": "${match.value}"`);
                    
                    if (!keyTranslations.has(match.key)) {
                        keyTranslations.set(match.key, {});
                    }
                    const keyTranslation = keyTranslations.get(match.key);
                    
                    // 处理所有语言
                    Object.keys(langIndexes).forEach(lang => {
                        if (hasContent[lang]) {
                            const langValue = cleanTranslation(row[langIndexes[lang]]);
                            if (langValue) {
                                keyTranslation[lang] = langValue;
                            }
                        }
                    });
                } else {
                    // 多个匹配，尝试智能分割
                    console.log(`  多个匹配，尝试智能分割:`);
                    matchingKeys.forEach(match => {
                        console.log(`    - "${match.key}": "${match.value}"`);
                    });
                    
                    // 处理每种语言的翻译
                    Object.keys(langIndexes).forEach(lang => {
                        if (!hasContent[lang]) return;
                        
                        const langValue = cleanTranslation(row[langIndexes[lang]]);
                        if (!langValue) return;
                        
                        const keyTranslationMap = splitMultipleTranslations(enValue, langValue, matchingKeys);
                        
                        // 保存翻译结果
                        Object.entries(keyTranslationMap).forEach(([key, translation]) => {
                            if (!keyTranslations.has(key)) {
                                keyTranslations.set(key, {});
                            }
                            keyTranslations.get(key)[lang] = translation;
                        });
                    });
                }
            }
        }

        // 为每种语言生成翻译文件
        const translations = {};
        Object.keys(langIndexes).forEach(lang => {
            if (hasContent[lang]) {
                translations[lang] = {};
            }
        });

        // 处理每个JSON key
        Object.entries(baseJson).forEach(([key, enValue]) => {
            const keyTranslation = keyTranslations.get(key);
            
            Object.keys(langIndexes).forEach(lang => {
                if (!hasContent[lang]) return;
                
                if (keyTranslation && keyTranslation[lang]) {
                    translations[lang][key] = keyTranslation[lang];
                } else {
                    // 如果没有找到翻译，使用原始英文值
                    translations[lang][key] = enValue;
                }
            });
        });

        // 生成翻译文件
        Object.entries(translations).forEach(([lang, data]) => {
            const outputPath = path.join(outputDir, `${lang}.json`);
            fs.writeJsonSync(outputPath, data, { spaces: 2 });
            console.log(`已生成 ${fileBase}/${lang}.json`);
        });

        // 输出匹配统计
        console.log(`\n匹配统计:`);
        console.log(`- 总共处理了 ${Object.keys(baseJson).length} 个JSON key`);
        console.log(`- 找到翻译的key: ${keyTranslations.size} 个`);
        console.log(`- 未找到翻译的key: ${Object.keys(baseJson).length - keyTranslations.size} 个`);

        if (keyTranslations.size > 0) {
            console.log(`\n已找到翻译的key:`);
            for (const [key, translations] of keyTranslations) {
                const langCount = Object.keys(translations).length;
                console.log(`  - ${key} (${langCount} 种语言)`);
            }
        }

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