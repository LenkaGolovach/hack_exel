import React from "react";

const Toolbar = ({ onCreate, onSave, onOpen }) => {
  return (
    <div style={{ marginBottom: "10px", display: "flex", gap: "10px" }}>
      <button onClick={onCreate}>Создать таблицу</button>
      <button onClick={onSave}>Сохранить как .xlsx</button>
      <input
        type="file"
        onChange={(e) => onOpen(e.target.files[0])}
        accept=".xlsx"
        style={{ marginLeft: "10px" }}
      />
    </div>
  );
};

export default Toolbar;
