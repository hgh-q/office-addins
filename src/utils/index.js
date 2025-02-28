// 检测浏览器支持技术
// writeExcel("B4", `ReadableStream is not supported`)
// if (typeof EventSource !== 'undefined') {
//   writeExcel("A1", `EventSource is supported`)
// } else {
//   writeExcel("B1", `EventSource is not supported`)
// }

// if ("WebSocket" in window) {
//   writeExcel("A2", `WebSocket is supported`)
// } else {
//   writeExcel("B2", `WebSocket is not supported`)
// }

export function parseDeepSeekBoxedResult(boxedString) {
    // 检查输入是否为 null 或 undefined
    if (boxedString == null) {
        return null;
    }
    let match = null

    // 使用正则表达式提取 \boxed{} 内的内容
    const boxedContentRegex = /\\boxed\{([^}]*)\}/;
    match = boxedString.match(boxedContentRegex);

    // 如果没有匹配到 \boxed{} 格式的内容
    // writeExcel("B1", 1)
    if (!match) {
        try {
            const excelFormulaRegex = /=.*\(.+\)/g;
            match = content.match(excelFormulaRegex);
            // writeExcel("B2", 125235)
        } catch (error) {
            // writeExcel("B9", JSON.stringify(error))
        }
    }

    // writeExcel("B3", 1)
    if (!match) {
        return boxedString
    }
    // writeExcel("B4", JSON.stringify(match))
    // 提取的内容
    const content = match[1];

    // 尝试将内容转换为数值
    let numericValue;
    try {
        numericValue = parseFloat(content);
        if (!isNaN(numericValue)) {
            return numericValue;
        }
    } catch (error) {
        // 转换失败，返回原始内容
        return content;
    }

    // 如果转换失败，返回原始内容
    return content;
}
