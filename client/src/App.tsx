import React from 'react';

function App() {
  // in BG validate data structure

  return (
    <div className="wrapper">
      <div className="flex-row flex-center header">
        <h1>Export your data to Excel spreadsheet</h1>
      </div>
      <div className="container mx-auto">
        <div className="flex-row flex-center">
          <div className="flex-col">
            <div className="flex-row">Select data structure</div>
            <div className="flex-row">
              <select id="data-structure" name="dataStructure">
                <option value="json"> JSON </option>
                <option value="xml"> XML </option>
                <option value="csv"> CSV </option>
              </select>
            </div>
            <div className="flex-row">Paste text / Drag & Drop</div>
            <div className="flex-row">
              <input type="text" />
            </div>
            <div className="flex-row">
              <button>Download</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
