import React, { useState } from "react";
import Header from './components/Header';
import InputBox from './components/InputBox';
import Dialogue from './components/Dialogue';
import 'whatwg-fetch';
import "./taskpane.css";
// import { writeExcel } from "../../utils/excel";

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

  const sendMessageChat = async () => {
    if (!message.trim()) return;

    setLoading(true);
    setMessage('');
    setMessages(prevMessages => [...prevMessages, { content: message, role: "user" }]);

    try {
      const apiUrl = "http://127.0.0.1:5000/api/chat";
      // const apiUrl = "https://api.siliconflow.cn/v1/chat/completions";
      // const apiKey = "sk-ooshywirgmrcdismctrllimnudbctvhhzybuzbqipervbrjy";

      // const requestBody = {
      //   "model": "deepseek-ai/DeepSeek-R1-Distill-Llama-70B",
      //   "messages": [{ "role": "user", "content": "test" }],
      // };

      // const response = await fetch(apiUrl, {
      //   method: "POST",
      //   headers: {
      //     "Content-Type": "application/json",
      //     "Authorization": Bearer ${apiKey},
      //   },
      //   body: JSON.stringify(requestBody),
      // });
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message }),
      });

      const responseBody = await response.json();

      if (response.ok && response.status === 200) {
        const answer = responseBody?.answer || "No answer returned.";
        setMessages(prevMessages => [
          ...prevMessages,
          { content: answer, role: "bot" }
        ]);

        // Excel update part
        await Excel.run(async (context) => {
          let sheet = context.workbook.worksheets.getActiveWorksheet();
          let range = sheet.getRange("A1");
          range.values = [[answer]];
          await context.sync();
        });
      } else {
        handleError(`请求失败状态: ${response.status}`);
      }
    } catch (error) {
      handleError(`请求失败：${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const getDSMessages = (messages) => {
    console.log(23, messages.filter(item=> ["user", "assistant"].includes(item.role)).map(({role, content})=>({role, content})))
    return messages.filter(item=> ["user", "assistant"].includes(item.role)).map(({role, content})=>({role, content}))
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

  const fetchMessages  = async (DSMessages) => {
    console.log(`DSMessages：${DSMessages}`)
    try {
      const apiUrl = "http://127.0.0.1:5000/api/stream";
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: DSMessages }),
      });
      if (!response.ok) {
        handleError(`请求失败：${response.status}, ${response.statusText}`);
        return;
      }
      const isStream = response.headers.get("Content-Type")?.includes("text/event-stream")
      if (isStream){
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let done = false;
        let reasoningContent = ''; // 存储 reasoning_content
        let content = ''; // 存储 content
        const messageId = Date.now(); // 或者你可以用其他方式生成唯一 ID，例如自增计数器


        setMessages(prevMessages => [...prevMessages, {id:messageId, content: "", role: "bot" }, {id:messageId, content: "", role: "assistant" }]);


        while (!done) {
          const { value, done: doneReading } = await reader.read();
          done = doneReading;
          let decodedString = decoder.decode(value, { stream: true });

          // 移除 'data: ' 前缀
          decodedString = decodedString.replace(/^data: /, '');
        
          // 解析 JSON
          try {
            const jsonData = JSON.parse(decodedString);
            if (jsonData.choices[0].delta.reasoning_content){
              reasoningContent += jsonData.choices[0].delta.reasoning_content;
              setMessages(prevMessages => {
                const updatedMessages = [...prevMessages];
                const botMessageIndex = updatedMessages.findIndex(msg => msg.id === messageId && msg.role === 'bot');
                if (botMessageIndex !== -1) {
                  updatedMessages[botMessageIndex].content = reasoningContent; // 更新bot的文本
                }
                return updatedMessages;
              });          
            }else{
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
      }else{
        const responseBody = await response.json();
        console.log(responseBody);
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
      test
    </div>
  );
};

export default App;
