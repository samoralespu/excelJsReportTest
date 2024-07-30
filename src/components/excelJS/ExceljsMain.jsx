import React, { useEffect, useState } from "react";
import { exportExcelFile } from "./ExportExcelFile";

const ExceljsMain = () => {
  const [data, setData] = useState([]);

  useEffect(() => {
    fetch("/message.json")
      .then((res) => res.json())
      .then(async (data) => {
        console.log("data: ", data);
        setData(data);
      })
      .then((json) => console.log(json));
  }, []);

  if (!data || data.length == 0) {
    return <>No data</>;
  }

  return (
    <div style={{ padding: "30px" }}>
      <button
        className="btn btn-primary float-end mt-2 mb-2"
        onClick={() => {
          exportExcelFile(data);
          console.log("onClick");
        }}
      >
        Export
      </button>
    </div>
  );
};

export default ExceljsMain;
