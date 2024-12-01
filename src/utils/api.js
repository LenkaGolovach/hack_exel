import axios from "axios";

const BASE_URL = "http://127.0.0.1:8000/api"; // Адрес вашего сервера FastAPI

export const createTable = async () => {
  const response = await axios.post(`${BASE_URL}/create`);
  return response.data.table;
};

export const saveTable = async (table) => {
  await axios.post(`${BASE_URL}/save`, { table });
};

export const openTable = async (file) => {
  const formData = new FormData();
  formData.append("file", file);

  const response = await axios.post(`${BASE_URL}/open`, formData, {
    headers: { "Content-Type": "multipart/form-data" },
  });
  return response.data.table;
};

export const saveFileToPath = async (file, savePath) => {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("save_path", savePath);
  
    const response = await axios.post(`${BASE_URL}/save-with-path`, formData);
    return response.data;
  };