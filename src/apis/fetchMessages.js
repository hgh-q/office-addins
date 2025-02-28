import { parseDeepSeekBoxedResult } from "@/utils/index"
const handleError = (errorText) => {
    setMessages(prevMessages => [
        ...prevMessages,
        { content: `${errorText}`, role: "bot" }
    ]);
};
export const fetchMessages = async (DSMessages, setMessages, setLoading, setAIResult) => {
    try {
        // if ('ReadableStream' in window && 'getReader' in ReadableStream.prototype) {
        if (false) {
            // const apiUrl = `${process.env.REACT_APP_API_HTTPS_URL}/api/stream`;
            const apiUrl = `https://127.0.0.1:5000/api/stream`;
            const response = await fetch(apiUrl, {
                method: 'POST',
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ messages: DSMessages }),
            })

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
                let reasoningContent = '';
                let content = '';
                const messageId = Date.now();

                setMessages(prevMessages => [...prevMessages, { id: messageId, content: "", role: "bot" }, { id: messageId, content: "", role: "assistant" }]);

                while (!done) {
                    const { value, done: doneReading } = await reader.read();
                    done = doneReading;
                    let decodedString = decoder.decode(value, { stream: true });

                    decodedString = decodedString.replace(/^data: /, '');
                    if (done) {
                        console.log(1232)
                        setLoading(false);
                        setAIResult(parseDeepSeekBoxedResult(content))
                    }

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
                        // setLoading(false);
                        // console.error(`解析JSON失败decodedString：${decodedString}:${error}`);
                    }
                }
            } else {
                const responseBody = await response.json();
                setLoading(false);
                setAIResult(parseDeepSeekBoxedResult(responseBody))
            }
        } else {
            let reasoningContent = '';
            let content = '';
            const messageId = Date.now();

            // const socket = new WebSocket(`${process.env.REACT_APP_API_WSS_URL}`);
            const socket = new WebSocket(`wss://127.0.0.1:5000`);

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
                    console.log("end")
                    setLoading(false);
                    setAIResult(parseDeepSeekBoxedResult(content))
                } else {
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
                console.log("关闭")
            };
        }
    } catch (error) {
        handleError(`请求失败：${error.message}`);
    }
}