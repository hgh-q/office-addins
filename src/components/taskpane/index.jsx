import React, { useState, useEffect } from "react";
import { fetchMessages } from "@/apis/fetchMessages"
import Header from '@/components/Header';
import Dialogue from '@/components/Dialogue';
import InputBox from '@/components/InputBox';
import 'whatwg-fetch';
import "./index.css";
import { readExcel, readUseExcel, writeExcel, writeSelectedRange, writeNonExcel, openMessageBox, openMyDialog } from "@/utils/excel";
import 'whatwg-fetch';

const App = ({ systemRoleDesc }) => {
    const [message, setMessage] = useState("");
    const [AIResult, setAIResult] = useState("");
    const [messages, setMessages] = useState([systemRoleDesc]);
    const [loading, setLoading] = useState(false);

    const isWrite = (val) => {
        if (val === 1) {
            writeSelectedRange(AIResult)
            setAIResult("")
        } else if (val === 0) {
            setAIResult("")
        }
    }

    const getDSMessages = (messages) => {
        return messages.filter(item => ["system", "user", "assistant"].includes(item.role)).map(({ role, content }) => ({ role, content }))
    }

    const sendMessageStream = async () => {
        if (!message.trim()) return;
        let ExcelData = []

        setLoading(true);
        setMessage('');
        try {
            ExcelData = await readUseExcel()
        } catch {
            ExcelData = [["项目", "价格"], ["a", 1], ["b", 2], ["c", 3], ["d", 4]]
        }
        if (!ExcelData) {
            throw new Error("读取Excel数据失败")
        }
        setMessages(prevMessages => {
            const updatedMessages = [...prevMessages, { content: `以下是Excel数据：${JSON.stringify(ExcelData)}。请完成用户需求：${message}。`, role: "user", text: message }];
            const DSMessages = getDSMessages(updatedMessages)
            fetchMessages(DSMessages, setMessages, setLoading, setAIResult)
            return updatedMessages;
        });
    };

    return (
        <div className="container_01">
            <Header />
            <Dialogue messages={messages} AIResult={AIResult} loading={loading} isWrite={isWrite} />
            <InputBox message={message} setMessage={setMessage} sendMessage={sendMessageStream} loading={loading} />
        </div>
    );
};

export default App;
