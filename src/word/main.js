import React from "react";
import "core-js"
import ReactDOM from "react-dom/client";
import Index from ".//index";

Office.onReady((info) => {
  const rootEle = document.getElementById("root")
  const root = ReactDOM.createRoot(rootEle);
  root.render(<Index />);
});