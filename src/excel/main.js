import React from "react";
import "core-js"
import ReactDOM from "react-dom/client";
import Index from ".//index";

Office.onReady((info) => {
  // if (info.host === Office.HostType.Excel) {
  //   // 在此处使用 Excel API
  //   console.log(Excel);
  // }
  const rootEle = document.getElementById("root")
  const root = ReactDOM.createRoot(rootEle);
  root.render(<Index />);
});