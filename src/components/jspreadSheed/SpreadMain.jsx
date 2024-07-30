import React, { useRef, useEffect } from "react";
import jspreadsheet from "jspreadsheet-ce";
import "../../../node_modules/jspreadsheet-ce/dist/jspreadsheet.css";

const SpreadMain = () => {
  const jRef = useRef(null);
  const options = {
    data: [[]],
    minDimensions: [10, 10],
  };

  useEffect(() => {
    if (!jRef.current.jspreadsheet) {
      jspreadsheet(jRef.current, options);
    }
  }, [options]);

  const addRow = () => {
    jRef.current.jexcel.insertRow();
  };

  return (
    <div>
      <div ref={jRef} />
      <br />
      <input type="button" onClick={addRow} value="Add new row" />
    </div>
  );
};

export default SpreadMain;
