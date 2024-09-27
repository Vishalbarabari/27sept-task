import React, { useState, useEffect } from 'react';
import ExcelReader from 'react-excel-reader';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import amqp from 'amqplib';

interface ExcelData {
  employee: string;
  activityCode: string;
}

function App() {
  const [data, setData] = useState<ExcelData[]>([]);
  const [selectedActivityCodes, setSelectedActivityCodes] = useState<string[]>([]);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const connectToRabbitMQ = async () => {
      try {
        const connection = await amqp.connect("amqps://guest:guest@localhost/rabbit"); 
        const channel = await connection.createChannel();
        const queue = 'RabbitM'; 

        channel.assertQueue(queue, { durable: true });

        channel.consume(queue, (msg) => {
          if (msg !== null) {
          
            const messageData = JSON.parse(msg.content.toString());
            setData((prevData) => [...prevData, messageData]);
          }
        });
      } catch (error) {
        console.error('Error connecting to RabbitMQ:', error);
      }
    };

    connectToRabbitMQ();
  }, []);

  const handleExcelFile = (data: ExcelData[]) => {
    const processedData = data.map((row) => ({
      employee: row[1],
      activityCode: row[2],
    }));

    setData(processedData);
    setError(null);
  };

  const handleActivityCodeChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setSelectedActivityCodes(event.target.value.split(','));
  };

  const handleGenerateReport = () => {
    const filteredData = data.filter((item) =>
      selectedActivityCodes.includes(item.activityCode)
    );

    const newExcelData = [
      ['Employee', 'Activity Code'],
      ...filteredData.map((item) => [item.employee, item.activityCode]),
    ];

    const ws = XLSX.utils.aoa_to_sheet(newExcelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append(wb, 'Sheet1', ws);

    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'   
 });

    saveAs(blob,   
 'report.xlsx');
  };

  return (
    <div>
      <h1>Report Generator</h1>
      <ExcelReader
        className="excel-reader"
        onFileLoaded={handleExcelFile}
        handleError={(e) => setError(e.message)}
      />
      {error && <p className="error">{error}</p>}
      <label htmlFor="activityCodes">Select Activity Codes:</label>
      <input
        type="text"
        id="activityCodes"
        value={selectedActivityCodes.join(',')}
        onChange={handleActivityCodeChange}
      />
      <button onClick={handleGenerateReport}>Generate Report</button>
    </div>
  );
}

export default App;