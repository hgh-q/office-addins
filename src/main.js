import React from "react";
import ReactDOM from "react-dom/client";
import Taskpane from "./pages/taskpane";
import "./main.css"
import { writeExcel } from "./utils/excel";

Office.onReady((info) => {
  console.log(`info.host：${info.host}`, `Office.HostType.Excel：${Office.HostType.Excel}`,)
  // if (info.host === Office.HostType.Excel) {
  //   // 在此处使用 Excel API
  //   console.log(Excel);
  // }
  const rootEle = document.getElementById("root")
  const root = ReactDOM.createRoot(rootEle);
  root.render(<Taskpane />);

  // if (!window.ReadableStream) {
  //   import('@stardazed/streams-polyfill')
  //     .then(() => {
  //       root.render(<Taskpane />);
  //     })
  //     .catch(err => {
  //       writeExcel("A3", `Polyfill loading failed${err}`)
  //     })
  // } else {
  //   root.render(<Taskpane />);
  // }
});

