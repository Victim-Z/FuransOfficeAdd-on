// 等待Office加载完成后初始化
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // 给按钮绑定点击事件
        document.getElementById('setMiaoColor').onclick = setMiaoBackgroundColor;
        document.getElementById('setCenterAlign').onclick = setCenterAlignment;
        document.getElementById('clearFormat').onclick = clearCellFormat;
        
        // 提示加载完成（可选）
        console.log("🐱 喵式格式化助手加载完成啦！");
    }
});

/**
 * 功能1：一键设置喵星蓝背景（#BDE0FE 清新浅蓝色）
 */
async function setMiaoBackgroundColor() {
    try {
        // 启动Excel操作
        await Excel.run(async (context) => {
            // 获取用户选中的单元格区域
            const range = context.workbook.getSelectedRange();
            // 设置背景色为喵星蓝
            range.format.fill.color = "#BDE0FE";
            // 执行操作
            await context.sync();
            alert("🐱 喵星蓝背景设置成功！");
        });
    } catch (error) {
        // 异常处理（避免崩溃）
        console.error("设置背景色失败：", error);
        alert("😿 设置失败啦！请先选中单元格再试试～");
    }
}

/**
 * 功能2：一键居中对齐（水平+垂直都居中）
 */
async function setCenterAlignment() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            // 水平居中
            range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
            // 垂直居中
            range.format.verticalAlignment = Excel.VerticalAlignment.center;
            await context.sync();
            alert("🐱 居中对齐设置成功！");
        });
    } catch (error) {
        console.error("设置居中失败：", error);
        alert("😿 设置失败啦！请先选中单元格再试试～");
    }
}

/**
 * 功能3：一键清除格式
 */
async function clearCellFormat() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            // 清除选中区域的所有格式
            range.format.clear();
            await context.sync();
            alert("🧹 格式清除成功！");
        });
    } catch (error) {
        console.error("清除格式失败：", error);
        alert("😿 清除失败啦！请先选中单元格再试试～");
    }
}