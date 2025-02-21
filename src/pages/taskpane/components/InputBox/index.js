import React, { useState } from "react";
import "./index.css"
import send48Png from "../../../../assets/images/send-48.png"
// import styled from "styled-components"

// const FooterCss = styled.footer`
//     background-color: #f1f1f1;
//     display: flex;
//     align-items: center;
//     justify-content: space-between;
// `;
// const InputBoxCss = styled.div`
//     display: flex;
//     flex-direction: column;
//     width: 100%;
// `;
// const TextareaCss = styled.textarea`
//     width: 100%;
//     height: 50px;
//     padding: 8px;
//     border-radius: 8px;
//     border: 1px solid #ddd;
//     resize: none;
// `;
// const ButtonCss = styled.button`
//     padding: 10px;
//     background-color: #28a745;
//     color: white;
//     border: none;
//     border-radius: 8px;
//     cursor: pointer;
//     transition: background-color 0.3s ease;

//     &:hover {
//       background-color: #218838;
//     }
// `;


const InputBox = ({ message, setMessage, sendMessage, loading }) => {
  
const handleKeyDown = (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    sendMessage()
  }
};

  return (
    <footer className={"footer"}>
      <div className={"input-box"} style={{background: loading ? "#f0f0f0" : ""}}>
        <textarea
          placeholder="请输入您的问题..."
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={loading}
        />
        <div className={"buttom-container"}>
          {!loading && <img onClick={sendMessage} src={send48Png} />}
        </div>
      </div>
    </footer>
  );
};

export default InputBox;
