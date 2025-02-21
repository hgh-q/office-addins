import React, { useState } from 'react';
import "./index.css"
import downArrow48 from "../../../../assets/images/down-arrow-48.png"
import upArrow48 from "../../../../assets/images/up-arrow-48.png"
// import styled from "styled-components"

// const ContainerCss = styled.div`
//   width: 100%;
//   margin: 0 auto;
//   padding: 10px;
//   display: flex;
//   flex-direction: column;
//   height: 500px;
//   background-color: #f7f7f7;
// `;

// const MessageListCss = styled.div`
//   width: 100%;
//   margin: 0 auto;
//   padding: 10px;
//   display: flex;
//   flex-direction: column;
//   overflow-y: auto;
//   flex-grow: 1;
// `;

// const UserMessageCss = styled.div`
//   background-color: #a1d4f4;
//   padding: 8px;
//   margin: 4px 0;
//   border-radius: 10px;
//   max-width: 60%;
//   align-self: flex-end;
// `;

// const BotMessageCss = styled.div`
//   background-color: #e0e0e0;
//   padding: 8px;
//   margin: 4px 0;
//   border-radius: 10px;
//   max-width: 60%;
//   align-self: flex-start;
// `;

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

const ChatComponent = ({ messages, AIResult, loading, setUserVerify }) => {
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
                <div className={"content"}>{`是否将结果："${AIResult}" 插入到选中区域`}</div>
                <div className={"btn-list"}>
                  <button onClick={() => setUserVerify(1)}>确认</button>
                  <button onClick={() => setUserVerify(0)}>取消</button>
                </div>
              </div>
            </div>]
        }
      </div>
    </div>
    // <ContainerCss>
    //   <MessageListCss>
    //     {messages.map((msg, index) =>{
    //       const MessageComponent = role === 'user' ? UserMessageCss : BotMessageCss;
    //       return <MessageComponent key={index}>{msg.content}</MessageComponent>;
    //     })}
    //   </MessageListCss>
    // </ContainerCss>
  );
}

export default ChatComponent;
