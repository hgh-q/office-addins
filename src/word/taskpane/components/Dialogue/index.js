import React, { useState } from 'react';
import "./index.css"
import downArrow48 from "@/assets/images/down-arrow-48.png"
import upArrow48 from "@/assets/images/up-arrow-48.png"

const Message = ({ msg, index }) => {
  const { role, text, content } = msg
  const [isExpanded, setIsExpanded] = useState(true); // 控制每条消息的展开状态

  const toggleExpand = () => {
    if (role === "bot") {
      setIsExpanded(!isExpanded);
    }
  };

  return (
    <div className="message-container" key={index}>
      <div className={`message ${role === 'user' ? "user-message" : role === 'bot' ? "bot-message" : "bot-thinking"}`} onClick={toggleExpand} style={{ cursor: role === "bot" ? "pointer" : "default" }}>
        {
          role === "bot" && <div className="Thinking">Thinking
            <img src={isExpanded ? upArrow48 : downArrow48} />
          </div>
        }
        {isExpanded ? text || content : ""}
      </div>
    </div>
  );
};

const ChatComponent = ({ messages, AIResult, loading, isWrite }) => {
  const lastMessage = messages[messages.length - 1]
  return (

    <div className={"container"}>
      <div className={"message-list"}>
        {messages.map((msg, index) => (
          <Message msg={msg} index={index} key={index} />
        ))}
        {
          !loading && lastMessage.role === "assistant" && AIResult !== "" && [
            <div className={"message-container"}>
              <div className={"user-verify message"}>
                <div className="content">
                  {`是否将结果：`}
                  <code className="highlight">{AIResult}</code>
                  {` 插入到选中区域`}
                </div>
                <div className={"btn-list"}>
                  <button onClick={() => isWrite(1)}>确认</button>
                  <button onClick={() => isWrite(0)}>取消</button>
                </div>
              </div>
            </div>]
        }
      </div>
    </div>
  );
}

export default ChatComponent;
