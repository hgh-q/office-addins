import React from "react";
import "core-js"
import ReactDOM from "react-dom/client";
import Taskpane from "./taskpane/index";
import { readExcel, readUseExcel, writeExcel, writeSelectedRange, writeNonExcel, openMessageBox, openMyDialog } from "@/utils/excel";

Office.onReady((info) => {
  // if (info.host === Office.HostType.Excel) {
  //   // 在此处使用 Excel API
  //   console.log(Excel);
  // }
  const rootEle = document.getElementById("root")
  const root = ReactDOM.createRoot(rootEle);
  root.render(<Taskpane />);
});