import React from "react";
import ReactDOM from "react-dom/client";

Office.onReady((info) => {
  const rootEle = document.getElementById("root")
  const root = ReactDOM.createRoot(rootEle);
  root.render(<div>word</div>);
});