import React from "react"
import Taskpane from "@/components/taskpane"

export default () => {
    const systemRoleDesc = { "role": "system", "content": `你是一名Word分析师，我会为你提供word内容，需要根据我的需求返回结果、公式或代码`, text: "你是一名Word分析师" }

    return <>
        <div style={{ width: "100%", height: "100%" }}>
            <Taskpane systemRoleDesc={systemRoleDesc}></Taskpane>
        </div>
    </>
}