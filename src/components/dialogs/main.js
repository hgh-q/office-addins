import React from "react";
import "core-js"
import ReactDOM from "react-dom/client";
import Popup from "./popup.jsx";

Office.onReady((info) => {
  const rootEle = document.getElementById("root")
  const root = ReactDOM.createRoot(rootEle);
  root.render(<Popup />);
});