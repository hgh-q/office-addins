import React from "react"
import Taskpane from "@/components/taskpane"

export default () => {
    const systemRoleDesc = { "role": "system", "content": `身份：你是一名Excel数据分析师；需求：我会为你提供二维数组excel表格内容，需要根据我的需求返回结果；返回内容要求1.将最终答案写入boxed{}中2.LaTeX 表格改为标准表格`, text: "你是一名Excel数据分析师" }

    return <>
        <div style={{ width: "100%", height: "100%" }}>
            <Taskpane systemRoleDesc={systemRoleDesc}></Taskpane>
        </div>
    </>
}