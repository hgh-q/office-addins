import React, { useState } from "react";
import Header from './components/Header';
import Dialogue from './components/Dialogue';
import InputBox from './components/InputBox';
import 'whatwg-fetch';
import "./taskpane.css";
import { writeExcel } from "../../utils/excel";
import 'whatwg-fetch';
// import 'web-streams-polyfill';
// import 'streams-polyfill';
// import "@stardazed/streams-polyfill"

const App = () => {

  const [message, setMessage] = useState("");
  const [messages, setMessages] = useState([]);
  const [loading, setLoading] = useState(false);

  const handleError = (errorText) => {
    setMessages(prevMessages => [
      ...prevMessages,
      { content: `${errorText}`, role: "bot" }
    ]);
  };

  const getDSMessages = (messages) => {
    return messages.filter(item => ["user", "assistant"].includes(item.role)).map(({ role, content }) => ({ role, content }))
  }

  const sendMessageStream = async () => {
    if (!message.trim()) return;

    setLoading(true);
    setMessage('');
    setMessages(prevMessages => {
      const updatedMessages = [...prevMessages, { content: message, role: "user" }];
      const DSMessages = getDSMessages(updatedMessages)
      fetchMessages(DSMessages)
      return updatedMessages;
    });
  };

  const fetchMessages = async (DSMessages) => {
    try {
      // if ('ReadableStream' in window && 'getReader' in ReadableStream.prototype) {
      if (false) {
        const apiUrl = "http://127.0.0.1:5000/api/stream";
        const response = await fetch(apiUrl, {
          method: 'POST',
          headers: { "Content-Type": "application/json" },
          // mode: "cors",  // 显式指定为 CORS 请求
          body: JSON.stringify({ messages: DSMessages }),
        })
        console.log(`response：${JSON.stringify(response)}`)
        if (!response.ok) {
          handleError(`请求失败：${response.status}, ${response.statusText}`);
          return;
        }

        if (!response.body) {
          throw new Error('No body in response');
        }
        const isStream = response.headers.get("Content-Type")?.includes("text/event-stream")
        if (isStream) {
          const reader = response.body.getReader();
          const decoder = new TextDecoder();
          let done = false;
          let reasoningContent = ''; // 存储 reasoning_content
          let content = ''; // 存储 content
          const messageId = Date.now(); // 或者你可以用其他方式生成唯一 ID，例如自增计数器


          setMessages(prevMessages => [...prevMessages, { id: messageId, content: "", role: "bot" }, { id: messageId, content: "", role: "assistant" }]);


          while (!done) {
            const { value, done: doneReading } = await reader.read();
            done = doneReading;
            let decodedString = decoder.decode(value, { stream: true });

            // 移除 'data: ' 前缀
            decodedString = decodedString.replace(/^data: /, '');

            // 解析 JSON
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
            } catch (error) {
              // console.error(`解析JSON失败decodedString：${decodedString}:${error}`);
            }
          }
        } else {
          const responseBody = await response.json();
          console.log(responseBody);
        }
      } else {
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
        try {
          writeExcel("E1", 1)
        } catch {

        }
        const socket = new WebSocket('wss://localhost:5000');  // WebSocket 连接到后端服务器
        try {
          writeExcel("E2", 2)
        } catch {

        }
        socket.onopen = function (event) {
          console.log('WebSocket connection opened.');

          // 发送请求到服务器，使用 WebSocket 发送消息
          const messageData = { messages: DSMessages };
          socket.send(JSON.stringify(messageData));
        };

        socket.onmessage = function (event) {
          const data = event.data;
          console.log('Received data:', data);
          writeExcel("D1", data)

          if (data === '[END]') {
            console.log('Stream ended');
          } else {
            // 处理从服务器接收到的数据
            try {
              const jsonData = JSON.parse(data);
              // 更新 UI 或做其他处理
            } catch (error) {
              console.error('Failed to parse message:', data);
            }
          }
        };

        socket.onerror = function (event) {
          try {
            writeExcel("F1", `WebSocket error：${JSON.stringify(event)}`)
          } catch {

          }
          console.error('WebSocket error:', event);
        };

        socket.onclose = function (event) {
          try {
            writeExcel("F2", `WebSocket error：${JSON.stringify(event)}`)
          } catch {

          }
          console.log('WebSocket connection closed:', event);
        };
      }
    } catch (error) {
      handleError(`请求失败：${error.message}`);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="container">
      {/* <Header />
      <Dialogue messages={messages} />
      <InputBox message={message} setMessage={setMessage} sendMessage={sendMessageStream} loading={loading} /> */}
      <Header />
      <Dialogue messages={messages} />
      <InputBox message={message} setMessage={setMessage} sendMessage={sendMessageStream} loading={loading} />
    </div>
  );
};

export default App;
