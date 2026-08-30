import "./storagePolyfill.js"; // must load before App.jsx uses window.storage
import React from "react";
import ReactDOM from "react-dom/client";
import SurgicalCaseLog from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <SurgicalCaseLog />
  </React.StrictMode>
);
