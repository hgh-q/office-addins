import React, { useState } from "react";
import "./index.css"
import send48Png from "@/assets/images/send-48.png"

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
