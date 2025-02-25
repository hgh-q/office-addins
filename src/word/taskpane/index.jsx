import React, { useState, useEffect } from "react";
import Header from './components/Header';
import Dialogue from './components/Dialogue';
import InputBox from './components/InputBox';
import 'whatwg-fetch';
import "./taskpane.css";
import { readExcel, readUseExcel, writeExcel, writeSelectedRange, writeNonExcel, openMessageBox, openMyDialog } from "@/utils/excel";
import 'whatwg-fetch';

const App = () => {
  const [message, setMessage] = useState("");
  const [AIResult, setAIResult] = useState("");
  const [messages, setMessages] = useState([
    { "role": "system", "content": `你是一名Word分析师，我会为你提供word内容，需要根据我的需求返回结果、公式或代码`, text: "你是一名Word分析师" },
  ]);
  const [loading, setLoading] = useState(false);

  const isWrite = (val) => {
    if (val === 1) {
      writeSelectedRange(AIResult)
      setAIResult("")
    } else if (val === 0) {
      setAIResult("")
    }
  }

  const handleError = (errorText) => {
    setMessages(prevMessages => [
      ...prevMessages,
      { content: `${errorText}`, role: "bot" }
    ]);
  };


  function parseDeepSeekBoxedResult(boxedString) {
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
      // ExcelData = await readExcel("A1:I35")
    } catch {
      ExcelData = [["项目", "价格"], ["a", 1], ["b", 2], ["c", 3], ["d", 4]]
    }
    if (!ExcelData) {
      throw new Error("读取Excel数据失败")
    }
    setMessages(prevMessages => {
      const updatedMessages = [...prevMessages, { content: `以下是Excel数据：${JSON.stringify(ExcelData)}。请完成用户需求：${message}。`, role: "user", text: message }];
      const DSMessages = getDSMessages(updatedMessages)
      fetchMessages(DSMessages)
      return updatedMessages;
    });
  };

  const fetchMessages = async (DSMessages) => {
    try {
      let reasoningContent = ''; // 存储 reasoning_content
      let content = ''; // 存储 content
      const messageId = Date.now(); // 或者你可以用其他方式生成唯一 ID，例如自增计数器

      // const socket = new WebSocket(`${process.env.REACT_APP_API_WSS_URL}`);  // WebSocket 连接到后端服务器
      const socket = new WebSocket(`wss://127.0.0.1:5000`);  // WebSocket 连接到后端服务器

      socket.onopen = function (event) {
        console.log('WebSocket connection opened.');
        setMessages(prevMessages => [...prevMessages, { id: messageId, content: "", role: "bot" }, { id: messageId, content: "", role: "assistant" }]);
        const messageData = { messages: DSMessages };
        socket.send(JSON.stringify(messageData));
      };

      socket.onmessage = function (event) {
        const data = event.data;
        const decodedString = data.replace(/^data: /, '');
        if (decodedString === '[END]') {
          try {
            setLoading(false);
            setAIResult(parseDeepSeekBoxedResult(content))
          } catch {
            console.log('Stream ended');
          }
        } else {
          // 处理从服务器接收到的数据
          try {
            const jsonData = JSON.parse(decodedString);
            if (jsonData.choices[0].delta.reasoning_content) {
              reasoningContent += jsonData.choices[0].delta.reasoning_content;
              setMessages(prevMessages => {
                const updatedMessages = [...prevMessages];
                const botMessageIndex = updatedMessages.findIndex(msg => msg.id === messageId && msg.role === 'bot');
                if (botMessageIndex !== -1) {
                  updatedMessages[botMessageIndex].content = reasoningContent; // 更新bot的文本
                }
                return updatedMessages;
              });
            } else {
              const assistantContent = jsonData.choices[0].delta.content
              if (assistantContent !== null) {
                content += jsonData.choices[0].delta.content;
                // 更新 assistant 消息内容
                setMessages(prevMessages => {
                  const updatedMessages = [...prevMessages];
                  const assistantMessageIndex = updatedMessages.findIndex(msg => msg.id === messageId && msg.role === 'assistant');
                  if (assistantMessageIndex !== -1) {
                    updatedMessages[assistantMessageIndex].content = content; // 更新assistant的文本
                  }
                  return updatedMessages;
                });
              }
            }
          } catch (error) {
            // console.error(`解析JSON失败decodedString：${decodedString}:${error}`);
          }
        }
      };

      socket.onerror = function (event) {
        try {
          // TODO: excel弹窗报错
        } catch {
          console.error('WebSocket error:', event);
        }
      };

      socket.onclose = function (event) {
      };
    } catch (error) {
      handleError(`请求失败：${error.message}`);
    }
  }

  return (
    <div className="container_01">
      <Header />
      <Dialogue messages={messages} AIResult={AIResult} loading={loading} isWrite={isWrite} />
      <InputBox message={message} setMessage={setMessage} sendMessage={sendMessageStream} loading={loading} />
    </div>
  );
};

export default App;
