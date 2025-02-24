import React, { useEffect } from "react";
import "./index.css"
import { readExcel, readUseExcel, writeExcel, writeSelectedRange, writeNonExcel, openMessageBox, openMyDialog } from "@/utils/excel";

export default () => {
    useEffect(() => {
        document.title = "身份证部分变星号"
        document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
        const checkboxes = document.querySelectorAll('input[type="checkbox"]');

        checkboxes.forEach(checkbox => {
            checkbox.addEventListener('change', () => {
                // 如果当前选中了一个复选框，则取消其他复选框的选中状态
                checkboxes.forEach(box => {
                    if (box !== checkbox) {
                        box.checked = false;
                    }
                });
            });
        });

        function sendStringToParentPage() {
            Office.context.ui.messageParent("userName");
        }

        async function tryCatch(callback) {
            try {
                await callback();
            } catch (error) {
                // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
                console.error(error);
            }
        }

        return () => {
            console.log('组件卸载');
        };
    }, []); // 空依赖数组，表示只在组件挂载时执行

    return (<div class="container">
        <div class="input-group">
            <label for="selectedRange">选择的数据区域：</label>
            <input type="text" id="selectedRange" readonly/>
        </div>
        <div class="input-group">
            
            <label for="birthDateOptions">变星选项：</label>
            <select id="birthDateOptions">
                <option value="showAll">隐藏出生日期(显示前6后4位)</option>
                <option value="hideFront4">保留前四位</option>
                <option value="hideBack6">隐藏后6位</option>
            </select>
        </div>
        <button id="ok-button" class="ms-Button">OK</button>
    </div>)
}