import React, { useState } from "react";
import "./taskpane.css";

const App = () => {
  const [message, setMessage] = useState("");
  const [response, setResponse] = useState("");

  const sendMessage = async () => {
    if (!message.trim()) return;

    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        range.format.fill.color = "yellow";
        await context.sync();
        setResponse(`单元格 ${range.address} 已高亮`);
      });
    } catch (error) {
      console.error(error);
      setResponse("执行失败，请检查控制台");
    }
  };

  return (
    <div className="container">
      <header className="header">
        <h1>Zone-OfficeAI</h1>
      </header>
      <section className="content">
        <p>在 Excel 里选择一个单元格，然后点击“发送”</p>
      </section>
      <footer className="footer">
        <div className="input-box">
          <textarea
            placeholder="请输入您的问题..."
            value={message}
            onChange={(e) => setMessage(e.target.value)}
          />
          <button className="send-btn" onClick={sendMessage}>
            发送
          </button>
        </div>
        {response && <p className="response">{response}</p>}
      </footer>
    </div>
  );
};

export default App;
