import React, { useState } from 'react';
import "./index.css"
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

const ChatComponent = ({ messages, AIResult, loading, setUserVerify }) => {
  console.log("messages：", messages)
  const lastMessage = messages[messages.length - 1]
  return (

    <div className={"container"}>
      <div className={"message-list"}>
        {messages.map((msg, index) => {
          return <div className={"message-container"}>
            <div className={`message ${msg.role === 'user' ? "user-message" : "bot-message"}`} key={index}>{msg?.text || msg.content}</div>
          </div>;
        })}
        {
          !loading && lastMessage.role === "assistant" && AIResult!=="" && [
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
    //       const MessageComponent = msg.role === 'user' ? UserMessageCss : BotMessageCss;
    //       return <MessageComponent key={index}>{msg.content}</MessageComponent>;
    //     })}
    //   </MessageListCss>
    // </ContainerCss>
  );
};

export default ChatComponent;
