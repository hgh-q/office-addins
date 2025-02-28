

function parseDeepSeekBoxedResult(boxedString) {
    if (boxedString == null) {
        return null;
    }

    let content = null;

    // 提取 \boxed{} 内的内容
    const boxedContentRegex = /\\boxed\{([^}]*)\}/;
    let match = boxedString.match(boxedContentRegex);
    console.log(match)
    if (match) {
        content = match[1].trim();
    } else {
        // 尝试匹配 Excel 公式
        const excelFormulaRegex = /=.*\(.+\)/g;
        match = boxedString.match(excelFormulaRegex);
        if (match) {
            content = match[0].trim();
        } else {
            return boxedString;
        }
    }
    // 判断是否是 LaTeX 表格
    if (/\\begin\{array\}/.test(content)) {
        return parseLatexTable(content);
    }

    // 尝试解析数值
    const numericValue = parseFloat(content);
    return isNaN(numericValue) ? content : numericValue;
}

// 解析 LaTeX 表格
function parseLatexTable(latexString) {
    const rows = latexString.split("\\hline").map(row => row.trim()).filter(row => row);
    let table = [];

    for (let row of rows) {
        // 移除 LaTeX 控制符，按 `&` 拆分单元格
        let cleanedRow = row.replace(/\\[a-zA-Z]+\{.*?\}/g, "").replace(/\\\\/g, "").trim();
        let cells = cleanedRow.split("&").map(cell => cell.trim());
        table.push(cells);
    }
    return table;
}


parseDeepSeekBoxedResult("为了将数据按照年龄排序，我们可以按照以下步骤进行： 1. 首先，我们有以下原始数据： | 姓名 | 性别 | 年龄 | 薪资 | |------|------|------|------| | a | 男 | 34 | 100 | | b | 男 | 53 | 200 | | c | 男 | 35 | 300 | | d | 女 | 36 | 400 | | e | 女 | 27 | 500 | | f | 男 | 53 | 600 | | g | 男 | 24 | 700 | | h | 女 | 34 | 800 | | i | 男 | 43 | 900 | | j | 女 | 37 | 1000 | | | | | 2800 | 2. 按照年龄从小到大排序后的结果如下： | 姓名 | 性别 | 年龄 | 薪资 | |------|------|------|------| | g | 男 | 24 | 700 | | e | 女 | 27 | 500 | | a | 男 | 34 | 100 | | h | 女 | 34 | 800 | | c | 男 | 35 | 300 | | d | 女 | 36 | 400 | | j | 女 | 37 | 1000 | | i | 男 | 43 | 900 | | b | 男 | 53 | 200 | | f | 男 | 53 | 600 | | | | | 2800 | ### 最终答案 \boxed{ \[ \begin{array}{|c|c|c|c|} \hline 姓名 & 性别 & 年龄 & 薪资 \\ \hline g & 男 & 24 & 700 \\ \hline e & 女 & 27 & 500 \\ \hline a & 男 & 34 & 100 \\ \hline h & 女 & 34 & 800 \\ \hline c & 男 & 35 & 300 \\ \hline d & 女 & 36 & 400 \\ \hline j & 女 & 37 & 1000 \\ \hline i & 男 & 43 & 900 \\ \hline b & 男 & 53 & 200 \\ \hline f & 男 & 53 & 600 \\ \hline & & & 2800 \\ \hline \end{array} \] }")