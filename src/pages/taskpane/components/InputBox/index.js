import React, { useState } from "react";
import "./index.css"
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
  return (
    <footer className={"footer"}>
      <div className={"input-box"}>
        <textarea
          placeholder="请输入您的问题..."
          value={message}
          onChange={(e) => setMessage(e.target.value)}
        />
        <button
          onClick={sendMessage}
          disabled={loading}
        >
          {loading ? '处理中...' : '发送'}
        </button>
      </div>
    </footer>
    // <FooterCss>
    //   <InputBoxCss>
    //     <TextareaCss
    //       placeholder="请输入您的问题..."
    //       value={message}
    //       onChange={(e) => setMessage(e.target.value)}
    //     />
    //     <ButtonCss
    //       onClick={sendMessage}
    //       disabled={loading}
    //     >
    //       {loading ? '处理中...' : '发送'}
    //     </ButtonCss>
    //   </InputBoxCss>
    // </FooterCss>
  );
};

export default InputBox;
