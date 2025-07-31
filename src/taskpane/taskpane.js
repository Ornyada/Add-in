Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    //There is a button to create it manually because in case user want to re-create the Masterfile without importing new file
    document.getElementById("importFolderBtn").addEventListener("click", async () => {
      const files = document.getElementById("folderInput").files;
      if (!files.length) {
        alert("Please select a folder first.");
        return;
      }
      const formData = new FormData();
      for (const file of files) {
        formData.append("files", file, file.webkitRelativePath);
      }
      await importFolder(formData);
    });
    const importDatalogBtn = document.getElementById("importDatalogBtn");
    if (importDatalogBtn) {
      importDatalogBtn.addEventListener("click", importFile);
    }
    //Select all button
    document.getElementById("compareAll").addEventListener("click", async () => {
      const checkboxes = document.querySelectorAll("#checkboxForm input[type='checkbox']");
      checkboxes.forEach((checkbox) => {
        checkbox.checked = !checkbox.checked;
      });
    });

    document.getElementById("compare").addEventListener("click", async () => {
      const checkboxes = document.querySelectorAll("#checkboxForm input[type='checkbox']");

      const UncheckedNames = Array.from(checkboxes)
        .filter((cb) => !cb.checked)
        .map((cb) => cb.value);
      console.log(UncheckedNames);

      const checkedNames = Array.from(checkboxes)
        .filter((cb) => cb.checked)
        .map((cb) => cb.value);
      console.log(checkedNames);

      await checkboxHide(UncheckedNames, checkedNames);
    });
  }
});
//convert stdf => xlsx
export async function run() {
  try {
    document.body.style.cursor = "wait";
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const baseName = "Masterfile";
      let sheetName = baseName;
      const existingNames = sheets.items.map((s) => s.name);
      // check all sheets name
      if (existingNames.includes(sheetName)) {
        //I'll be back!!!!
      } else {
        const newSheet = sheets.add(sheetName);
        const headers = ["Suite name", "Test name", "Test number", "Lsl_typ", "Usl_typ", "Units"];
        const headerRange = newSheet.getRangeByIndexes(0, 0, 1, headers.length); //determine the range of cells to input headers , index เริ่มนับที่ 0
        headerRange.values = [headers]; //input headers into cells
        headerRange.format.fill.color = "#43a0ec"; // Background of headers
        headerRange.format.font.bold = true;
        const sheet = context.workbook.worksheets.getItem("Masterfile");
        sheet.position = 0;
        sheet.activate();
      }
      await context.sync();
      document.body.style.cursor = "default";
    });
  } catch (error) {
    console.error("Error:", error);
    logToConsole("Error");
  }
}

//For importing datalog files
async function importFile() {
  document.body.style.cursor = "wait";
  const fileInput = document.getElementById("fileInput");
  const files = fileInput.files;
  const fileArray = Array.from(files);
  if (!files || files.length === 0) return;
  console.log("Amount of  file: %d", fileArray.length);
  logToConsole("Amount of  file: %d", fileArray.length);
  let file_processed = 0;
  for (const file of fileArray) {
    const isCSV = file.name.toLowerCase().endsWith(".csv");
    const isXLSX = file.name.toLowerCase().endsWith(".xlsx");
    const isSTDF = file.name.toLowerCase().endsWith(".stdf");
    try {
      if (isCSV || isXLSX) {
        console.log("file CSV or XLSX is processing");
        logToConsole("file CSV or XLSX is processing");
        //seperate converted datalog and limit files

        const reader = new FileReader();
        reader.onload = async function (e) {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetCount = workbook.SheetNames.length;
          if (sheetCount > 1) {
            file_processed = await uploadSelfConvertedDatalog(file, file_processed);
            logToConsole("Processed file = %d", file_processed);
          } else {
            logToConsole("EY datalog is importing");
            file_processed = await uploadEYdatalog(file, file_processed);
            logToConsole("Processed file = %d", file_processed);
          }
        };
        reader.readAsArrayBuffer(file);
        // display file name and path
        const importedList = document.getElementById("importedFilesList");
        if (importedList) {
          const listItem = document.createElement("li");
          listItem.textContent = `${file.name} - ${file.webkitRelativePath || file.name}`;
          importedList.appendChild(listItem);
        }
      } else if (isSTDF) {
        console.log("File is STDF");
        logToConsole("File is STDF");
        const formData = new FormData();
        if (!file) {
          console.warn("No file");
          return;
        }
        formData.append("files", file);
        console.log(`Processing: ${file.name}`);
        logToConsole(`Processing: ${file.name}`);
        document.body.style.cursor = "wait";
        const response = await fetch("https://limit-project-demo.onrender.com/upload-stdf/", {
          method: "POST",
          body: formData,
        });
        if (!response.ok) {
          const errorText = await response.text();
          console.error("STDF upload failed:", errorText);
          logToConsole("STDF upload failed:", errorText);
          return;
        }
        // import file as blob
        const blob = await response.blob();
        const downloadUrl = window.URL.createObjectURL(blob);

        // use orginal file.name and change end of file name to .xlsx
        const originalName = file.name.replace(/\.[^/.]+$/, "");
        const downloadName = `${originalName}.xlsx`;
        const a = document.createElement("a");
        a.href = downloadUrl;
        a.download = downloadName;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(downloadUrl);
        logToConsole("STDF converted and downloaded successfully");
      } else {
        console.warn(`Doesn't support ${file.name}`);
        logToConsole(`Doesn't support ${file.name}`);
      }
    } catch (err) {
      console.error(`Error while processing file: ${file.name}`, err);
      logToConsole(`Error while processing file: ${file.name}`);
    } finally {
      //write file name in InputFiles Sheet
      await Excel.run(async (context) => {
        let sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        let sheetName = "InputFiles";
        let existingNames = sheets.items.map((s) => s.name);
        let sheet;
        if (existingNames.includes(sheetName)) {
          sheet = sheets.getItem(sheetName);
        } else {
          sheet = sheets.add(sheetName);
          const headers = ["File_Name"];
          const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
          headerRange.values = [headers];
          headerRange.format.fill.color = "#C6EFCE";
          headerRange.format.font.bold = true;
          sheet.position = 0;
        }
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();
        const nextRow = usedRange.rowCount;
        const targetCell = sheet.getRangeByIndexes(nextRow, 0, 1, 1);
        targetCell.values = [[file.name]];
        await context.sync();
      });
    }
  }
  document.body.style.cursor = "default";
}
//For processing datalog that is converted by this tool
async function uploadSelfConvertedDatalog(file, file_processed) {
  document.body.style.cursor = "wait";
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      const formData = new FormData();
      formData.append("file", file);
      console.log(`Uploading Excel Datalog to API: ${file.name}`);
      logToConsole(`Uploading Excel Datalog to API: ${file.name}`);
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const mirSheet = workbook.Sheets["mir"];
      Excel.run(async (context) => {
        const masterSheet = context.workbook.worksheets.getItem("Masterfile");
        let usedRange;
        usedRange = masterSheet.getUsedRange();
        usedRange.load(["rowCount", "columnCount"]);
        await context.sync();
        const sheet = context.workbook.worksheets.getItem("Masterfile");
        const chunkSize = 1000;
        const totalRows = usedRange.rowCount;
        const totalCols = usedRange.columnCount;
        let allValues = [];
        for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
          const rowCount = Math.min(chunkSize, totalRows - startRow);
          const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
          range.load("values");
          await context.sync();
          allValues = allValues.concat(range.values);
        }
        let headers = allValues[0];
        console.log("headers: ", headers);

        //Insert JOB_NAM as a Product Name
        const mirData = XLSX.utils.sheet_to_json(mirSheet, { defval: "" });
        const productName = mirData[0]?.["JOB_NAM"]?.trim();
        let stagename = mirData[0]?.["TEST_COD"]?.trim();
        let productColIndex = headers.indexOf(productName);
        console.log("productColINdex before add product name or stage: %d", productColIndex);
        logToConsole("productColINdex before add product name or stage: %d", productColIndex);
        let Allproduct_stage = [];
        let StartStageCol;
        let EndStageCol;
        let allstagescount;
        for (let i = 0; i <= headers.length; i++) {
          if (headers[i] === "Can remove (Y/N)") {
            StartStageCol = i;
          }
          if (headers[i] === "Lsl_typ") {
            EndStageCol = i;
          }
        }
        allstagescount = EndStageCol - StartStageCol - 1;
        let temp;
        if (allstagescount > 0) {
          for (let i = StartStageCol + 1; i < EndStageCol; i++) {
            const Procell = headers[i];
            const stageCell = allValues[1][i];
            if (Procell && Procell.trim() !== "") {
              Allproduct_stage.push({
                name: Procell.trim(),
                stage: stageCell,
              });
              temp = Procell.trim();
            } else {
              Allproduct_stage.push({
                name: temp,
                stage: stageCell,
              });
            }
          }
        }

        // If there is no same product name then insert it
        if (productColIndex === -1) {
          const sheet = context.workbook.worksheets.getItem("Masterfile");
          const columnToInsert = sheet.getRange("F:F");
          columnToInsert.insert(Excel.InsertShiftDirection.right);
          const product_name_head = sheet.getRange("F1:F1");
          product_name_head.values = [[productName]];
          let Canremove_index = headers.indexOf("Can remove (Y/N)");
          if (Canremove_index < 0) {
            logToConsole("Can't find Can remove col");
            return;
          }
          let Lsl_typ_index = headers.indexOf("Lsl_typ");
          if (Lsl_typ_index < 0) {
            logToConsole("Can't find Lsl_typ col");
            return;
          }
          let Product_count = 0;
          for (let i = Canremove_index; i < Lsl_typ_index; i++) {
            let cell = usedRange.getCell(0, i);
            if (!isNaN(cell) || cell !== "") {
              Product_count++;
            }
          }
          const colors = ["#C6EFCE", "#FFEB9C", "#FFC7CE", "#e6cdfa"];
          const color = colors[Product_count % 4];
          product_name_head.format.fill.color = color;
          //add stage
          const stage_name_head = sheet.getRange("F2:F2");
          stage_name_head.values = [[stagename]];
          await context.sync();
        } else {
          //if product name is same then check if the stage is same
          const sheet = context.workbook.worksheets.getItem("Masterfile");
          await context.sync();
          const startCol = productColIndex;
          const stage_count = Allproduct_stage.filter((item) => item.name === productName).length; //how many stages does this product have
          console.log("product : %s , stage count : %d stage", productName, stage_count);
          let columnToInsert = sheet.getRangeByIndexes(0, startCol + 1, 1, 1);
          columnToInsert.insert(Excel.InsertShiftDirection.right);
          usedRange = sheet.getUsedRange();
          usedRange.load("rowCount");
          await context.sync();
          const stagerow = usedRange.rowCount;
          let stage_name_head = sheet.getRangeByIndexes(1, startCol + stage_count, stagerow, 1);
          stage_name_head.insert(Excel.InsertShiftDirection.right);
          usedRange = sheet.getUsedRange();
          //usedRange.load("values");
          await context.sync();
          const stageCell = sheet.getCell(1, startCol + stage_count);
          stageCell.values = [[stagename]];
          console.log("startcol for merge: %d stagecount for merge: %d", startCol, stage_count);
          logToConsole("startcol for merge: %d stagecount for merge: %d", startCol, stage_count);
          const range = sheet.getRangeByIndexes(0, startCol, 1, stage_count);
          range.values = Array(1).fill(Array(stage_count).fill(productName));
          // fill productName in every cells of range Array(stage_count).fill(productName)
          // create sub array that has duplicate (productName) for (stage_count) times ex. ["ABC", "ABC", "ABC"] Array(1).fill(...) => array 2 dim (1 row n cols) → match to expected .values
          range.merge();
          await context.sync();
        }
        usedRange = masterSheet.getUsedRange();
        await context.sync();
        masterSheet.activate();
        console.log("Completely added product name and stage");
        logToConsole("Completely added product name and stage");
        return fetch("https://limit-project-demo.onrender.com/process-self-converted-datalog/", {
          method: "POST",
          body: formData,
        })
          .then((res) => res.json())
          .then((data) => {
            let TestData = data.test_data;
            if (TestData !== null) {
              logToConsole("process-datalog-excel fetched successfully");
            }
            Excel.run(async (context) => {
              const sheet = context.workbook.worksheets.getItem("Masterfile");
              let usedRange = sheet.getUsedRange();
              usedRange.load(["rowCount", "columnCount"]);
              await context.sync();

              let chunkSize = 1000;
              let totalRows = usedRange.rowCount;
              let totalCols = usedRange.columnCount;
              let allValues = [];
              for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
                const rowCount = Math.min(chunkSize, totalRows - startRow);
                const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
                range.load("values");
                await context.sync();

                allValues = allValues.concat(range.values);
              }
              let headers = allValues[0];

              let TestnumColIndex = headers.indexOf("Test number");
              const SuiteColIndex = headers.indexOf("Suite name");
              const TestColIndex = headers.indexOf("Test name");
              if (TestnumColIndex === -1) {
                console.error("Test number column is not found");
                logToConsole("Test number is not found");
                return;
              }
              if (SuiteColIndex === -1) {
                console.error("Suite name is not found");
                logToConsole("Suite name is not found");
                return;
              }
              if (TestColIndex === -1) {
                console.error("Test name is not found");
                logToConsole("Test name is not found");
                return;
              }

              const testNameRange = sheet.getRangeByIndexes(
                2,
                TestColIndex,
                allValues.length - 2,
                1
              );
              testNameRange.load("values");
              await context.sync();
              logToConsole("Determined Allcolindex and testNamerange");
              let existingTestNames = [];
              try {
                existingTestNames = testNameRange.values.flat().filter((v) => v !== "");
              } catch (err) {
                console.error("Error while processing testNameRange.values:", err);
                logToConsole("Error while processing testNameRange.values: %s", err.message || err);
                return;
              }
              if (!Array.isArray(TestData)) {
                console.error("TestData isn't an array or isn't downloaded");
                logToConsole("TestData isn't an array or isn't downloaded");
                return;
              }
              let newTests = [];
              try {
                newTests = TestData.filter((item) => !existingTestNames.includes(item.test_name));
              } catch (err) {
                console.error("Error while TestData.filter", err);
                logToConsole("Error while TestData.filter: %s", err.message || err);
                return;
              }
              if (!Array.isArray(allValues)) {
                console.error("allValues isn't an array");
                logToConsole("allValues isn't an array");
                return;
              }
              let startRow = allValues.length;
              let suiteRange, testRange;
              let suiteValues = [];
              let testValues = [];
              try {
                if (newTests.length > 0) {
                  const testNumbers = newTests.map((t) => [t?.test_number ?? ""]);
                  // fill test numbers
                  if (TestnumColIndex === -1) {
                    logToConsole("Can't find Test number column in header");
                    return;
                  }
                  const writeRange = sheet.getRangeByIndexes(
                    startRow,
                    TestnumColIndex,
                    newTests.length,
                    1
                  );
                  writeRange.values = testNumbers;
                  await context.sync();
                  // fill suite name and test name
                  suiteRange = sheet.getRangeByIndexes(startRow, SuiteColIndex, newTests.length, 1);
                  testRange = sheet.getRangeByIndexes(startRow, TestColIndex, newTests.length, 1);
                  suiteValues = newTests.map((t) => [t.suite_name]);
                  testValues = newTests.map((t) => [t.test_name]);
                  suiteRange.values = suiteValues;
                  testRange.values = testValues;
                  await context.sync();
                } else {
                  logToConsole("There's no new tests");
                }
              } catch (err) {
                console.error("Error while processing newTests:", err);
              }
              // read uploaded Excel to import product name
              const data = new Uint8Array(e.target.result);
              const workbook = XLSX.read(data, { type: "array" });
              const mirSheet = workbook.Sheets["mir"];
              const mirData = XLSX.utils.sheet_to_json(mirSheet, { defval: "" });
              const productName = mirData[0]?.["JOB_NAM"]?.trim();
              let productColIndex = headers.indexOf(productName);
              if (productColIndex === -1) {
                console.error("Can't find product name :", productName);
                logToConsole("Can't find product name :", productName);
                return;
              }
              Allproduct_stage.push({
                name: productName,
                stage: stagename,
              });
              let stage_count = Allproduct_stage.filter((item) => item.name === productName).length;
              let stage_array_index;
              let stage_range = sheet.getRangeByIndexes(1, productColIndex, 1, stage_count);
              stage_range.load("values");
              await context.sync();
              for (let i = 0; i <= stage_count; i++) {
                console.log("stage %d = %s", i, stage_range.values[0][i]);
                if (stage_range.values[0][i] === stagename) {
                  stage_array_index = i;
                  break;
                }
              }

              if (stage_array_index === undefined) {
                console.error("Can't find stage name index :", stagename);
                logToConsole("Can't find stage name index:", stagename);
              }
              console.log(
                "productColIndex: %d, stage_count: %d, stageArrayIndex: %d ",
                productColIndex,
                stage_count,
                stage_array_index
              );
              usedRange = sheet.getUsedRange();
              await context.sync();
              usedRange.load(["rowCount", "columnCount"]);
              await context.sync();
              totalRows = usedRange.rowCount;
              totalCols = usedRange.columnCount;
              allValues = [];
              for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
                const rowCount = Math.min(chunkSize, totalRows - startRow);
                const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
                range.load("values");
                await context.sync();
                allValues = allValues.concat(range.values);
              }
              headers = allValues[0];
              const testNameRangeAll = sheet.getRangeByIndexes(
                2,
                TestColIndex,
                allValues.length - 2,
                1
              );
              testNameRangeAll.load("values");
              await context.sync();
              // create YNValues by mapping test_name -> YN_check
              const allTestNames = testNameRangeAll.values.map((row) => row[0]);
              let YNValues = [];
              try {
                YNValues = allTestNames.map((testName) => {
                  const match = TestData.find((item) => item.test_name === testName);
                  return [match ? match.YN_check : ""];
                });
              } catch (err) {
                console.error("Error while creating YNValues:", err);
              }
              let YNRange = sheet.getRangeByIndexes(
                2,
                productColIndex + stage_array_index,
                YNValues.length,
                1
              );
              YNRange.load("values");
              await context.sync();

              if (YNValues.length === 0) {
                console.warn("No Y/N check data");
                logToConsole("No Y/N check data");
              } else {
                console.log("YN.length of %s %s is %d", productName, stagename, YNValues.length);
                logToConsole("YN.length of %s %s is %d", productName, stagename, YNValues.length);
              }
              YNRange.values = YNValues;
              await context.sync();
              const IsUsedIndex = headers.indexOf("Is used (Y/N)");
              let IsUsedDataRange = sheet.getRangeByIndexes(
                2,
                IsUsedIndex,
                YNRange.values.length,
                1
              );
              IsUsedDataRange.load("values");
              await context.sync();
              let IsUsedData = IsUsedDataRange.values;
              if (!Array.isArray(IsUsedData) || IsUsedData.length === 0) {
                IsUsedData = Array.from({ length: YNRange.values.length }, () => [""]);
              }
              //create IsUsed data
              for (let i = 0; i < YNRange.values.length; i++) {
                if (YNRange.values[i][0] === "Y") {
                  if (IsUsedData[i][0] === "Partial" || IsUsedData[i][0] === "No") {
                    IsUsedData[i][0] = "Partial";
                  } else if (IsUsedData[i][0] === "") {
                    IsUsedData[i][0] = "All";
                  }
                } else {
                  if (IsUsedData[i][0] === "All" || IsUsedData[i][0] === "Partial") {
                    IsUsedData[i][0] = "Partial";
                  } else IsUsedData[i][0] = "No";
                }
              }
              IsUsedDataRange.values = IsUsedData;
              await context.sync();
              //conditional formatting color
              const conditionalFormat = YNRange.conditionalFormats.add(
                Excel.ConditionalFormatType.containsText
              );
              conditionalFormat.textComparison.format.fill.color = "#C6EFCE";
              conditionalFormat.textComparison.rule = {
                operator: Excel.ConditionalTextOperator.contains,
                text: "Y",
              };
              const IsUsedkeywords = ["All", "Partial"];
              const colors = ["#C6EFCE", "#FFEB9C"];
              for (let i = IsUsedDataRange.conditionalFormats.count - 1; i >= 0; i--) {
                IsUsedDataRange.conditionalFormats.getItemAt(i).delete();
              }
              await context.sync();
              for (let i = 0; i < IsUsedkeywords.length; i++) {
                const word = IsUsedkeywords[i];
                const color = colors[i];

                const conditionalFormat = IsUsedDataRange.conditionalFormats.add(
                  Excel.ConditionalFormatType.containsText
                );
                conditionalFormat.textComparison.format.fill.color = color;
                conditionalFormat.textComparison.rule = {
                  operator: Excel.ConditionalTextOperator.contains,
                  text: word,
                };
              }
              await context.sync();
              console.log("Finished processing one file");
              logToConsole("Finished processing one file");
              file_processed++;
              resolve(file_processed);
              document.body.style.cursor = "default";
            });
          });
      });
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
//For processing datalog from EY
async function uploadEYdatalog(file, file_processed) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async function (e) {
      const formData = new FormData();
      formData.append("file", file);
      console.log(`Uploading Excel Datalog to API: ${file.name}`);
      logToConsole(`Uploading Excel Datalog to API: ${file.name}`);
      const data = new Uint8Array(e.target.result);
      const response = await fetch("https://limit-project-demo.onrender.com/process-EY/", {
        method: "POST",
        body: formData,
      });
      /*{ result will look like this
          "data": [
          {
            "test_number": 61,
            "suite_name": "ivn_std_init",
            "test_name": "OtpMapCollabNetRev",
            "YN_check": "Y",
            "product": "BirdRock",
            "stage": "FH3"
          },*/
      if (!response.ok) {
        console.error("Upload failed");
        logToConsole("Upload failed");
        return;
      }
      const result = await response.json();
      const EYdata = result.data;
      //seperate difference stage data
      let All_EY_Stage_Product = [];
      let tempStage;
      let tempProductname;
      let Allproduct = [];
      for (const item of EYdata) {
        if (tempProductname !== item.product) {
          tempProductname = item.product;
          tempStage = item.stage;
          All_EY_Stage_Product.push({
            name: item.product,
            stage: item.stage,
          });
          Allproduct.push(tempProductname);
        } else if (item.stage !== tempStage) {
          tempStage = item.stage;
          All_EY_Stage_Product.push({
            name: item.product,
            stage: item.stage,
          });
        }
      }
      console.log("All product: ", Allproduct);
      console.log("All EY stage and product: ", All_EY_Stage_Product);
      //loop for all product
      for (let tempProductname of Allproduct) {
        let OneProduct_Allstage = All_EY_Stage_Product.filter(
          (item) => item.name === tempProductname
        );
        console.log("OneProduct_Allstage: ", OneProduct_Allstage);
        //loop for each stage of one product
        for (let item of OneProduct_Allstage) {
          let productName = item.name;
          let stageName = item.stage.toLowerCase();
          await Check_product_stage(productName, stageName);
          let OneStage_data = EYdata.filter(
            (content) => productName === content.product && item.stage === content.stage
          );
          console.log("OneStage_data: ", OneStage_data);
          await WriteNewtest(OneStage_data);
          await YN(OneStage_data, productName, stageName);
        }
      }

      logToConsole("Import EY file successfully");
      file_processed++;
      resolve(file_processed);
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
//For importing test program folder
async function importFolder(formData) {
  let arranged_stages = [];
  let Allpair = [];
  let Allfirst = [];
  let Alllast = [];
  let limit_compare = [];
  await Excel.run(async (context) => {
    document.body.style.cursor = "wait";
    const response = await fetch("https://limit-project-demo.onrender.com/upload-folder/", {
      method: "POST",
      body: formData,
    });
    if (!response.ok) {
      console.error("Upload failed");
      logToConsole("Upload failed");
      return;
    }
    logToConsole("Import Folder fetched successfully");
    const result = await response.json();
    const mfhFiles = result.mfh_files || [];
    //display .mfh name list in UI
    const mfhList = document.getElementById("mfh-list");
    mfhList.innerHTML = "";
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    //check and create InputFiles sheet
    let sheetName = "InputFiles";
    let existingNames = sheets.items.map((s) => s.name);
    let inputSheet;
    if (existingNames.includes(sheetName)) {
      inputSheet = sheets.getItem(sheetName);
    } else {
      inputSheet = sheets.add(sheetName);
      const headers = ["File_Name"];
      const headerRange = inputSheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.values = [headers];
      headerRange.format.fill.color = "#C6EFCE";
      headerRange.format.font.bold = true;
      inputSheet.position = 0;
      await context.sync();
    }

    //check and create Masterfile sheet
    sheetName = "Masterfile";
    let masterSheet;
    existingNames = sheets.items.map((s) => s.name);
    if (!existingNames.includes(sheetName)) {
      console.log("There is no Masterfile yet...Creating Masterfile");
      logToConsole("There is no Masterfile yet...Creating Masterfile");
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const baseName = "Masterfile";
      let sheetName = baseName;
      const existingNames = sheets.items.map((s) => s.name);
      if (existingNames.includes(sheetName)) {
        //I'll be back!!!!
      } else {
        masterSheet = sheets.add(sheetName);
        masterSheet.position = 0;
      }
      masterSheet.activate();
    }
    mfhFiles.forEach((fileName) => {
      const li = document.createElement("li");
      li.textContent = `${fileName} (ready)`;
      li.style.cursor = "pointer";
      li.addEventListener("click", async () => {
        // remove class 'selected-file' from other list
        document.querySelectorAll("li").forEach((item) => {
          item.classList.remove("selected-file");
        });
        // add new class for selected list
        li.classList.add("selected-file");
        const res = await fetch(
          `https://limit-project-demo.onrender.com/process-testtable/?filename=${encodeURIComponent(
            fileName
          )}`
        );
        if (!res.ok) {
          const container = document.getElementById("download-links");
          container.innerHTML = `<p style="color:red;">Failed to process ${fileName}</p>`;
          return;
        }
        li.textContent = fileName;
        const data = await res.json();
        displayResults(data.files);
        await Excel.run(async (context) => {
          document.body.style.cursor = "wait";
          const masterSheet = context.workbook.worksheets.getItem("Masterfile");
          let all_limit_stage = [];
          for (let file_index = 0; file_index < data.files.length; file_index++) {
            const usedRange = masterSheet.getUsedRange();
            usedRange.load(["values", "rowCount", "columnCount"]);
            await context.sync();
            let headers = usedRange.values[0] || [];
            let stages = usedRange.values[1] || [];
            await context.sync();
            let file = data.files[file_index];
            if (file.status === "ok" && Array.isArray(file.content)) {
              const fileHeaders = file.content[0];
              const stageHeaders = file.content[1];
              //write col from first uploaded file
              if (file_index === 0) {
                for (let col = 0; col < fileHeaders.length; col++) {
                  const fileheader = fileHeaders[col];
                  const stageheader = stageHeaders[col];
                  all_limit_stage.push({
                    name: fileheader,
                    stage: stageheader,
                  });
                  if (fileheader && fileheader !== (headers[col] || "")) {
                    masterSheet.getCell(0, col).values = [[fileheader]];
                    await context.sync();
                  }
                  if (stageheader && stageheader !== stages[col]) {
                    masterSheet.getCell(1, col).values = [[stageheader]];
                    await context.sync();
                  }
                }
                usedRange.load(["rowCount", "columnCount"]);
                await context.sync();
                const sheet = context.workbook.worksheets.getItem("Masterfile");
                const chunkSize = 1000;
                const totalRows = usedRange.rowCount;
                const totalCols = usedRange.columnCount;
                let allValues = [];
                for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
                  const rowCount = Math.min(chunkSize, totalRows - startRow);
                  const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
                  range.load("values");
                  await context.sync();
                  allValues = allValues.concat(range.values);
                }
                headers = allValues[0];
                stages = allValues[1];
              } else {
                usedRange.load(["rowCount", "columnCount"]);
                await context.sync();
                let UnitsIndex = headers.indexOf("Units");
                if (UnitsIndex === NaN) {
                  return;
                }
                //find new_stage
                for (let col = 0; col < headers.length; col++) {
                  const fileheader = fileHeaders[col];
                  const stageheader = stageHeaders[col];
                  let samestage;
                  if (fileheader === "Lsl" || fileheader === "Usl") {
                    samestage = all_limit_stage.find((item) => item.stage === stageheader);
                    if (samestage === undefined || samestage === "") {
                      logToConsole("new stage is %s", stageheader);
                      all_limit_stage.push({
                        name: fileheader,
                        stage: stageheader,
                      });
                      let newstageColRange = masterSheet.getRangeByIndexes(
                        0,
                        UnitsIndex + 3, //after spec col
                        usedRange.rowCount,
                        2
                      );
                      newstageColRange.insert(Excel.InsertShiftDirection.right);
                      logToConsole("Insert new col");
                      await context.sync();
                      usedRange.load(["rowCount", "columnCount"]);
                      await context.sync();
                      masterSheet.getCell(0, UnitsIndex + 3).values = [["Lsl"]];
                      masterSheet.getCell(0, UnitsIndex + 4).values = [["Usl"]];
                      masterSheet.getCell(1, UnitsIndex + 3).values = stageheader;
                      masterSheet.getCell(1, UnitsIndex + 4).values = stageheader;
                      await context.sync();
                    }
                  }
                }
              }
              await context.sync();
            }
          }
          await context.sync();
          let usedRange = masterSheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          let sheet = context.workbook.worksheets.getItem("Masterfile");
          let chunkSize = 1000;
          let totalRows = usedRange.rowCount;
          let totalCols = usedRange.columnCount;
          let allValues = [];
          for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
            const rowCount = Math.min(chunkSize, totalRows - startRow);
            const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
            range.load("values");
            await context.sync();
            allValues = allValues.concat(range.values);
          }
          let headers = allValues[0];
          let stages = allValues[1];
          let nextRow = usedRange.rowCount;
          await context.sync();
          //arrange stages
          let wafer_stage = [];
          let final_stage = [];
          let a_stage = [];
          let wh = [];
          let wr = [];
          let wc = [];
          let wi = [];
          let ww = [];
          let fh = [];
          let fr = [];
          let fc = [];
          let fi = [];
          let fw = [];
          usedRange = masterSheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          chunkSize = 1000;
          totalRows = usedRange.rowCount;
          totalCols = usedRange.columnCount;
          allValues = [];
          for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
            const rowCount = Math.min(chunkSize, totalRows - startRow);
            const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
            range.load("values");
            await context.sync();
            allValues = allValues.concat(range.values);
          }
          headers = allValues[0];
          stages = allValues[1];
          await context.sync();
          console.log(stages);
          wafer_stage = stages.filter((item) => item[0] === "w");
          final_stage = stages.filter((item) => item[0] === "f");
          a_stage = stages.filter((item) => item[0] === "a");
          console.log(wafer_stage);
          console.log(final_stage);
          console.log(a_stage);
          wh = wafer_stage
            .filter((item) => item[1] === "h")
            .sort((a, b) => {
              return parseInt(a[0].replace("wh", "")) - parseInt(b[0].replace("wh", ""));
            });
          wr = wafer_stage
            .filter((item) => item[1] === "r")
            .sort((a, b) => {
              return parseInt(a[0].replace("wr", "")) - parseInt(b[0].replace("wr", ""));
            });
          wc = wafer_stage
            .filter((item) => item[1] === "c")
            .sort((a, b) => {
              return parseInt(a[0].replace("wc", "")) - parseInt(b[0].replace("wc", ""));
            });
          wi = wafer_stage
            .filter((item) => item[1] === "i")
            .sort((a, b) => {
              return parseInt(a[0].replace("wi", "")) - parseInt(b[0].replace("wi", ""));
            });
          ww = wafer_stage
            .filter((item) => item[1] === "w")
            .sort((a, b) => {
              return parseInt(a[0].replace("ww", "")) - parseInt(b[0].replace("ww", ""));
            });
          console.log(wh);
          console.log(wr);
          console.log(wc);
          console.log(ww);
          console.log(wi);
          wafer_stage = [];
          if (wh.length !== 0) {
            wafer_stage.push(...wh);
          }
          if (wr.length !== 0) {
            wafer_stage.push(...wr);
          }
          if (wc.length !== 0) {
            wafer_stage.push(...wc);
          }
          if (ww.length !== 0) {
            wafer_stage.push(...ww);
          }
          if (wi.length !== 0) {
            wafer_stage.push(...wi);
          }
          arranged_stages.push(...wafer_stage);
          fh = final_stage
            .filter((item) => item[1] === "h")
            .sort((a, b) => {
              return parseInt(a[0].replace("fh", "")) - parseInt(b[0].replace("fh", ""));
            });
          fr = final_stage
            .filter((item) => item[1] === "r")
            .sort((a, b) => {
              return parseInt(a[0].replace("fr", "")) - parseInt(b[0].replace("fr", ""));
            });
          fc = final_stage
            .filter((item) => item[1] === "c")
            .sort((a, b) => {
              return parseInt(a[0].replace("fc", "")) - parseInt(b[0].replace("fc", ""));
            });
          fi = final_stage
            .filter((item) => item[1] === "i")
            .sort((a, b) => {
              return parseInt(a[0].replace("fi", "")) - parseInt(b[0].replace("fi", ""));
            });
          fw = final_stage
            .filter((item) => item[1] === "w")
            .sort((a, b) => {
              return parseInt(a[0].replace("fw", "")) - parseInt(b[0].replace("fw", ""));
            });
          console.log(fh);
          console.log(fr);
          console.log(fc);
          console.log(fw);
          console.log(fi);
          final_stage = [];
          if (fh.length !== 0) {
            final_stage.push(...fh);
          }
          if (fr.length !== 0) {
            final_stage.push(...fr);
          }
          if (fc.length !== 0) {
            final_stage.push(...fc);
          }
          if (fw.length !== 0) {
            final_stage.push(...fw);
          }
          if (fi.length !== 0) {
            final_stage.push(...fi);
          }
          arranged_stages.push(...final_stage);
          arranged_stages.push(...a_stage); // needs to fix this if there are more 'a' test
          //send stages data to checkbox
          const uniqueWh = wh.filter((_, index) => index % 2 === 0);
          const uniqueWr = wr.filter((_, index) => index % 2 === 0);
          const uniqueWc = wc.filter((_, index) => index % 2 === 0);
          const uniqueWi = wi.filter((_, index) => index % 2 === 0);
          const uniqueWw = ww.filter((_, index) => index % 2 === 0);
          const uniqueFh = fh.filter((_, index) => index % 2 === 0);
          const uniqueFr = fr.filter((_, index) => index % 2 === 0);
          const uniqueFc = fc.filter((_, index) => index % 2 === 0);
          const uniqueFi = fi.filter((_, index) => index % 2 === 0);
          const uniqueFw = fw.filter((_, index) => index % 2 === 0);
          const uniqueAr = a_stage.filter((_, index) => index % 2 === 0);
          let pairList = [];
          // match stages for comparison and prevent duplicate/self matching
          uniqueWh.forEach((w) => {
            uniqueFh.forEach((f) => {
              const pairId = `${w}__${f}`; // use __ to seperate stage
              const label = `${w} ? ${f}`;
              pairList.push({ id: pairId, label });
            });
          });
          uniqueWr.forEach((w) => {
            uniqueFr.forEach((f) => {
              const pairId = `${w}__${f}`;
              const label = `${w} ? ${f}`;
              pairList.push({ id: pairId, label });
            });
          });
          uniqueWc.forEach((w) => {
            uniqueFc.forEach((f) => {
              const pairId = `${w}__${f}`;
              const label = `${w} ? ${f}`;
              pairList.push({ id: pairId, label });
            });
          });
          uniqueWi.forEach((w) => {
            uniqueFi.forEach((f) => {
              const pairId = `${w}__${f}`;
              const label = `${w} ? ${f}`;
              pairList.push({ id: pairId, label });
            });
          });
          uniqueWw.forEach((w) => {
            uniqueFw.forEach((f) => {
              const pairId = `${w}__${f}`;
              const label = `${w} ? ${f}`;
              pairList.push({ id: pairId, label });
            });
          });
          uniqueFr.forEach((w) => {
            uniqueAr.forEach((f) => {
              const pairId = `${w}__${f}`;
              const label = `${w} ? ${f}`;
              pairList.push({ id: pairId, label });
            });
          });
          uniqueWh.forEach((a, i) => {
            uniqueWh.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueWr.forEach((a, i) => {
            uniqueWr.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueWc.forEach((a, i) => {
            uniqueWc.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueWi.forEach((a, i) => {
            uniqueWi.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueWw.forEach((a, i) => {
            uniqueWw.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueFh.forEach((a, i) => {
            uniqueFh.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueFr.forEach((a, i) => {
            uniqueFr.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueFc.forEach((a, i) => {
            uniqueFc.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueFi.forEach((a, i) => {
            uniqueFi.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          uniqueFw.forEach((a, i) => {
            uniqueFw.forEach((b, j) => {
              if (i < j) {
                const pairId = `${a}__${b}`;
                const label = `${a} ? ${b}`;
                pairList.push({ id: pairId, label });
              }
            });
          });
          console.log(pairList);
          // create checkbox from pairList
          pairList.forEach((pair) => {
            const labelWrapper = document.createElement("label");
            labelWrapper.className = "label-item";
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.id = pair.id;
            checkbox.name = "stagePair";
            checkbox.value = pair.id;
            labelWrapper.appendChild(checkbox);
            labelWrapper.appendChild(document.createTextNode(` ${pair.label}`));
            labelList.appendChild(labelWrapper);
            //collect each pair in array
            let first = pair.label.slice(0, 3); //first three letters
            let last = pair.label.slice(-3); //last three letters
            Allfirst.push(...[first]);
            Alllast.push(...[last]);
            let pair_header = [
              "LL " + first.toUpperCase() + " ? " + last.toUpperCase(),
              "UL " + first.toUpperCase() + " ? " + last.toUpperCase(),
            ];
            Allpair.push(...pair_header);
          });
          console.log("Allpair: ", Allpair);
          let SpecIndex = stages.indexOf("Spec");
          if (isNaN(SpecIndex)) {
            logToConsole("Can't find Spec column!");
            return;
          }
          logToConsole("Spec index is %d", SpecIndex);
          let arrange_range = masterSheet.getRangeByIndexes(
            1,
            SpecIndex + 2,
            1,
            arranged_stages.length
          );
          arrange_range.values = [arranged_stages];
          await context.sync();

          usedRange = masterSheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          chunkSize = 1000;
          totalRows = usedRange.rowCount;
          totalCols = usedRange.columnCount;
          allValues = [];
          for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
            const rowCount = Math.min(chunkSize, totalRows - startRow);
            const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
            range.load("values");
            await context.sync();
            allValues = allValues.concat(range.values);
          }
          headers = allValues[0];
          stages = allValues[1];
          nextRow = usedRange.rowCount;
          await context.sync();
          for (const file of data.files) {
            if (file.status === "ok" && Array.isArray(file.content)) {
              const fileHeaders = file.content[0];
              const stageHeaders = file.content[1];
              for (let i = 2; i < file.content.length; i++) {
                const row = file.content[i];
                const rowData = [];
                for (let col = 0; col < headers.length; col++) {
                  const header = headers[col];
                  if (header === "Lsl" || header === "Usl") {
                    const stageName = stages[col];
                    if (header === "Lsl") {
                      const file_stageIndex = stageHeaders.indexOf(stageName);
                      if (file_stageIndex === NaN) {
                        continue;
                      }
                      rowData.push(file_stageIndex !== -1 ? row[file_stageIndex] : "");
                    } else {
                      const file_stageIndex = stageHeaders.indexOf(stageName);
                      if (file_stageIndex === NaN) {
                        continue;
                      }
                      rowData.push(file_stageIndex !== -1 ? row[file_stageIndex + 1] : "");
                    }
                  } else {
                    const MasterheaderIndex = headers.indexOf(header);
                    const headerIndex = fileHeaders.indexOf(header);
                    if (MasterheaderIndex !== -1 && headerIndex !== -1) {
                      rowData[MasterheaderIndex] = row[headerIndex];
                    }
                  }
                }
                const targetRange = masterSheet.getRangeByIndexes(nextRow, 0, 1, headers.length);
                targetRange.values = [rowData];
                nextRow++;
              }
            }
          }
          await context.sync();
          //create columns for limit compare
          usedRange = masterSheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          chunkSize = 1000;
          totalRows = usedRange.rowCount;
          totalCols = usedRange.columnCount;
          allValues = [];
          for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
            const rowCount = Math.min(chunkSize, totalRows - startRow);
            const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
            range.load("values");
            await context.sync();
            allValues = allValues.concat(range.values);
          }
          headers = allValues[0];
          stages = allValues[1];
          nextRow = usedRange.rowCount;
          let Bin_s_index = headers.indexOf("Bin_s_num");
          await context.sync();
          masterSheet
            .getRangeByIndexes(0, Bin_s_index, usedRange.rowCount, 2)
            .insert(Excel.InsertShiftDirection.right);
          await context.sync();
          masterSheet.getCell(0, Bin_s_index).values = [["All LL ? Spec"]];
          masterSheet.getCell(0, Bin_s_index + 1).values = [["All UL ? Spec"]];
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          //insert new columns for limit compare
          masterSheet
            .getRangeByIndexes(0, Bin_s_index + 2, usedRange.rowCount, arranged_stages.length)
            .insert(Excel.InsertShiftDirection.right);
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          for (let col = SpecIndex + 2; col < SpecIndex + 2 + arranged_stages.length; col += 2) {
            limit_compare.push(...["LL Spec? " + stages[col], "UL Spec? " + stages[col]]);
          }
          console.log(limit_compare);
          console.log("limit_compare contains undefined?", limit_compare.includes(undefined));
          console.log("limit_compare contains null?", limit_compare.includes(null));
          masterSheet.getRangeByIndexes(0, Bin_s_index + 2, 1, limit_compare.length).values = [
            limit_compare,
          ];
          await context.sync();
          document.body.style.cursor = "default";
        });
        await Excel.run(async (context) => {
          await context.sync();
          await new Promise((resolve) => setTimeout(resolve, 100));
          document.body.style.cursor = "wait";
          //limit comparison
          const masterSheet = context.workbook.worksheets.getItem("Masterfile");
          const usedRange = masterSheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          let sheet = context.workbook.worksheets.getItem("Masterfile");
          let chunkSize = 1000;
          let totalRows = usedRange.rowCount;
          let totalCols = usedRange.columnCount;
          let allValues = [];
          for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
            const rowCount = Math.min(chunkSize, totalRows - startRow);
            const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
            range.load("values");
            await context.sync();
            allValues = allValues.concat(range.values);
          }
          let headers = allValues[0];
          let stages = allValues[1];
          let values = allValues;
          let rowCount = usedRange.rowCount;
          let columnCount = usedRange.columnCount;
          logToConsole("rowCount : %d columnCount : %d", rowCount, columnCount);
          let LLspecIndex = stages.indexOf("Spec");
          const All_LL_specIndex = headers.indexOf("All LL ? Spec");
          const All_UL_specIndex = headers.indexOf("All UL ? Spec");
          logToConsole("LLspecIndex : %d", LLspecIndex);
          let ULlastIndex = LLspecIndex + arranged_stages.length + 2;
          let limit = [];
          for (let r = 2; r < values.length; r++) {
            let rowData = [];
            for (let c = LLspecIndex; c <= ULlastIndex; c++) {
              rowData.push(values[r][c]); //collect all limits
            }
            limit.push(rowData);
          }
          let firstLLindex = headers.indexOf("All UL ? Spec") + 1;
          let All_LL_spec = [];
          let All_UL_spec = [];
          let ALL_compare_result = [];
          logToConsole("Limit length : %d", limit.length);
          for (let i = 0; i < limit.length; i++) {
            const row = limit[i];
            const specLL = row[0];
            const specUL = row[1];
            let in_outllResult = "In-spec";
            let in_outulResult = "In-spec";
            ALL_compare_result[i] = [];
            // start from index 2 because index 0,1 are specLL, specUL
            for (let j = 2; j + 1 < row.length; j += 2) {
              const LLvalue = row[j];
              const ULvalue = row[j + 1];
              if (LLvalue === undefined || ULvalue === undefined) {
                console.warn(`Missing LSL/USL at row ${i}, columns ${j} and ${j + 1}`);
                continue;
              }
              let LLspec_limResult = "";
              let ULspec_limResult = "";
              // LL comparison
              if (LLvalue !== "" && LLvalue != null && !isNaN(LLvalue)) {
                if (!(LLvalue >= specLL)) {
                  in_outllResult = "Out-spec";
                  LLspec_limResult = "Tighten";
                } else if (specLL < LLvalue) {
                  LLspec_limResult = "Widen";
                } else {
                  LLspec_limResult = "Same";
                }
              }
              // UL comparison
              if (ULvalue !== "" && ULvalue != null && !isNaN(ULvalue)) {
                if (!(ULvalue <= specUL)) {
                  in_outulResult = "Out-spec";
                  ULspec_limResult = "Tighten";
                } else if (specUL > ULvalue) {
                  ULspec_limResult = "Widen";
                } else {
                  ULspec_limResult = "Same";
                }
              }
              // Collect data of each row
              ALL_compare_result[i].push(LLspec_limResult);
              ALL_compare_result[i].push(ULspec_limResult);
            }
            All_LL_spec.push([in_outllResult]);
            All_UL_spec.push([in_outulResult]);
          }
          // Write Result in Excel
          masterSheet.getRangeByIndexes(
            2,
            firstLLindex,
            ALL_compare_result.length,
            ALL_compare_result[0].length
          ).values = ALL_compare_result;
          await context.sync();
          logToConsole("All_LL_spec length : %d", All_LL_spec.length);
          logToConsole("All_UL_spec length : %d", All_UL_spec.length);
          let All_LL_specRange = masterSheet.getRangeByIndexes(
            2,
            All_LL_specIndex,
            All_LL_spec.length,
            1
          );
          let All_UL_specRange = masterSheet.getRangeByIndexes(
            2,
            All_UL_specIndex,
            All_UL_spec.length,
            1
          );
          All_LL_specRange.values = All_LL_spec;
          All_UL_specRange.values = All_UL_spec;
          await context.sync();
          //Insert "Is used (Y/N)", "Can remove (Y/N)"
          const columnToInsert = masterSheet.getRangeByIndexes(0, 3, rowCount, 2);
          columnToInsert.insert(Excel.InsertShiftDirection.right);
          masterSheet.getCell(0, 3).values = "Is used (Y/N)";
          masterSheet.getCell(0, 4).values = "Can remove (Y/N)";
          await context.sync();
          logToConsole("Limit Compare Successed");
          //In/Out-spec conditional formatting
          const IN_OUTkeywords = ["Out-spec", "In-spec"];
          const colors = ["#ff9c9c", "#C6EFCE"];
          for (let i = 0; i < IN_OUTkeywords.length; i++) {
            const word = IN_OUTkeywords[i];
            const color = colors[i];
            const LL_conditionalFormat = All_LL_specRange.conditionalFormats.add(
              Excel.ConditionalFormatType.containsText
            );
            LL_conditionalFormat.textComparison.format.fill.color = color;
            LL_conditionalFormat.textComparison.rule = {
              operator: Excel.ConditionalTextOperator.contains,
              text: word,
            };
            const UL_conditionalFormat = All_UL_specRange.conditionalFormats.add(
              Excel.ConditionalFormatType.containsText
            );
            UL_conditionalFormat.textComparison.format.fill.color = color;
            UL_conditionalFormat.textComparison.rule = {
              operator: Excel.ConditionalTextOperator.contains,
              text: word,
            };
          }

          await context.sync();
          document.body.style.cursor = "default";
        });
        await Excel.run(async (context) => {
          document.body.style.cursor = "wait";
          await context.sync();
          //Insert limit vs limit col
          const masterSheet = context.workbook.worksheets.getItem("Masterfile");
          await context.sync();
          let usedRange = masterSheet.getUsedRange();
          usedRange.load(["rowCount", "columnCount"]);
          await context.sync();
          let sheet = context.workbook.worksheets.getItem("Masterfile");
          let chunkSize = 1000;
          let totalRows = usedRange.rowCount;
          let totalCols = usedRange.columnCount;
          let allValues = [];
          for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
            const rowCount = Math.min(chunkSize, totalRows - startRow);
            const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
            range.load("values");
            await context.sync();
            allValues = allValues.concat(range.values);
          }
          let headers = allValues[0];
          let stages = allValues[1];
          await context.sync();
          let Bin_s_index = headers.indexOf("Bin_s_num");
          logToConsole("Bin s index: %d", Bin_s_index);
          masterSheet
            .getRangeByIndexes(0, Bin_s_index, usedRange.rowCount, Allpair.length)
            .insert(Excel.InsertShiftDirection.right);
          masterSheet.getRangeByIndexes(0, Bin_s_index, 1, Allpair.length).values = [Allpair];
          await context.sync();
          let LLspecIndex = stages.indexOf("Spec");
          logToConsole("LLspecIndex : %d", LLspecIndex);
          let ULlastIndex = LLspecIndex + arranged_stages.length + 2;
          let limit = [];
          for (let r = 2; r < allValues.length; r++) {
            let rowData = [];
            for (let c = LLspecIndex; c <= ULlastIndex; c++) {
              rowData.push(allValues[r][c]); //collect all limits
            }
            limit.push(rowData);
          }
          //limit vs limit comparison
          let ALL_compare_result = [];
          for (let i = 0; i < Allfirst.length; i++) {
            const f = Allfirst[i];
            const l = Alllast[i];
            const fIndex = stages.indexOf(f);
            const lIndex = stages.indexOf(l);
            if (fIndex < 0 || lIndex < 0) {
              console.warn(`ไม่พบ stage: ${f} หรือ ${l}`);
              continue;
            }
            for (let r = 0; r < limit.length; r++) {
              const row = limit[r];
              const LLfirst = row[fIndex - LLspecIndex];
              const ULfirst = row[fIndex - LLspecIndex + 1];
              const LLlast = row[lIndex - LLspecIndex];
              const ULlast = row[lIndex - LLspecIndex + 1];
              let LLlim_limResult = "";
              let ULlim_limResult = "";

              if (LLfirst !== "" && LLfirst != null && !isNaN(LLfirst)) {
                if (LLlast !== "" && LLlast != null && !isNaN(LLlast)) {
                  if (LLfirst < LLlast) {
                    LLlim_limResult = "Widen";
                  } else if (LLfirst > LLlast) {
                    LLlim_limResult = "Tighten";
                  } else {
                    LLlim_limResult = "Same";
                  }
                }
              }
              if (ULfirst !== "" && ULfirst != null && !isNaN(ULfirst)) {
                if (ULlast !== "" && ULlast != null && !isNaN(ULlast)) {
                  if (ULfirst > ULlast) {
                    ULlim_limResult = "Widen";
                  } else if (ULfirst < ULlast) {
                    ULlim_limResult = "Tighten";
                  } else {
                    ULlim_limResult = "Same";
                  }
                }
              }
              if (!ALL_compare_result[r]) {
                ALL_compare_result[r] = [];
              }
              ALL_compare_result[r].push(LLlim_limResult);
              ALL_compare_result[r].push(ULlim_limResult);
            }
          }
          console.log("All compare:", ALL_compare_result);
          let lastSpecLimit_index = headers.indexOf(limit_compare[limit_compare.length - 1]);
          console.log(lastSpecLimit_index);
          if (lastSpecLimit_index < 0) {
            logToConsole("Can't find last spec vs limit col");
            return;
          }
          masterSheet.getRangeByIndexes(
            2,
            lastSpecLimit_index + 1,
            ALL_compare_result.length,
            ALL_compare_result[0].length
          ).values = ALL_compare_result;
          await context.sync();
          //fill header color
          usedRange.load("columnCount");
          await context.sync();
          for (let i = 0; i < usedRange.columnCount; i++) {
            masterSheet.getCell(0, i).format.fill.color = "#c4d4f3";
          }
          await context.sync();
          document.body.style.cursor = "default";
        });

        //write file name in InputFiles sheet
        await Excel.run(async (context) => {
          document.body.style.cursor = "wait";
          const inputSheet = context.workbook.worksheets.getItem("InputFiles");
          const usedRange = inputSheet.getUsedRange();
          usedRange.load("rowCount");
          await context.sync();
          const nextRow = usedRange.rowCount;
          const targetCell = inputSheet.getRangeByIndexes(nextRow, 0, 1, 1);
          targetCell.values = [[fileName]];
          await context.sync();
          logToConsole("Successfully limit files imported");
          document.body.style.cursor = "default";
        });
      });
      mfhList.appendChild(li);
    });
    document.body.style.cursor = "default";
  });
}

// display .mfh files
function displayResults(files) {
  document.body.style.cursor = "wait";
  const container = document.getElementById("download-links");
  container.innerHTML = ""; // empty container
  let startRow = 3;
  files.forEach((file) => {
    const rowCount = file.content?.length || 0;
    const endRow = rowCount - 3 + startRow;
    const div = document.createElement("div");
    div.innerHTML = `
      <p><b>${file.path}</b> - ${file.status} (${startRow}-${endRow})</p>
    `;
    container.appendChild(div);
    startRow = endRow + 1;
  });
  document.body.style.cursor = "default";
}

function logToConsole(format, ...args) {
  const consoleDiv = document.getElementById("consoleOutput");

  // create timestamp
  const now = new Date();
  const timestamp = now.toLocaleString(); // or use now.toISOString() for international standard format

  // replace each %s, %d, %f
  let formatted = format;
  let argIndex = 0;
  formatted = formatted.replace(/%[sdif]/g, (match) => {
    const arg = args[argIndex++];
    switch (match) {
      case "%d":
      case "%i":
        return parseInt(arg);
      case "%f":
        return parseFloat(arg).toFixed(2);
      case "%s":
      default:
        return String(arg);
    }
  });

  // create a message with timestamp
  const line = document.createElement("div");
  line.textContent = `[${timestamp}] > ${formatted}`;
  consoleDiv.appendChild(line);
  consoleDiv.scrollTop = consoleDiv.scrollHeight;
}

async function checkboxHide(UncheckedNames, checkedNames) {
  await Excel.run(async (context) => {
    document.body.style.cursor = "wait";
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    let masterSheet = context.workbook.worksheets.getItem("Masterfile");
    let usedRange = masterSheet.getUsedRange();
    await context.sync();
    usedRange.load(["rowCount", "columnCount"]);
    await context.sync();
    console.log(usedRange.rowCount);
    const chunkSize = 1000;
    const totalRows = usedRange.rowCount;
    const totalCols = usedRange.columnCount;
    let allValues = [];
    for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
      const rowCount = Math.min(chunkSize, totalRows - startRow);
      const range = masterSheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
      range.load("values");
      await context.sync();
      allValues = allValues.concat(range.values);
    }
    let headers = allValues[0];
    let LLpair_header;
    let LLpairIndex;

    for (const pair of UncheckedNames) {
      let first = pair.slice(0, 3); //first three letters
      let last = pair.slice(-3); //last three letters
      LLpair_header = "LL " + first.toUpperCase() + " ? " + last.toUpperCase();
      console.log(LLpair_header);
      LLpairIndex = headers.indexOf(LLpair_header);
      if (LLpairIndex === -1) {
        logToConsole("can't find an index to hide");
        return;
      }
      logToConsole("Hiding : %s ? %s", first.toUpperCase(), last.toUpperCase());
      try {
        masterSheet
          .getRangeByIndexes(0, LLpairIndex - 2, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = true;
        masterSheet
          .getRangeByIndexes(0, LLpairIndex - 1, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = true;
        masterSheet
          .getRangeByIndexes(0, LLpairIndex, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = true;
        masterSheet
          .getRangeByIndexes(0, LLpairIndex + 1, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = true;
        await context.sync();
      } catch (err) {
        console.log("Can't hide col due to: ", err);
      }
    }
    await context.sync();
    for (const pair of checkedNames) {
      let first = pair.slice(0, 3); //first three letters
      let last = pair.slice(-3); //last three letters
      LLpair_header = "LL " + first.toUpperCase() + " ? " + last.toUpperCase();
      LLpairIndex = headers.indexOf(LLpair_header);
      if (LLpairIndex === -1) {
        logToConsole("can't find an index to hide");
        return;
      }
      try {
        masterSheet
          .getRangeByIndexes(0, LLpairIndex - 2, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = false;
        masterSheet
          .getRangeByIndexes(0, LLpairIndex - 1, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = false;
        masterSheet
          .getRangeByIndexes(0, LLpairIndex, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = false;
        masterSheet
          .getRangeByIndexes(0, LLpairIndex + 1, usedRange.rowCount, 1)
          .getEntireColumn().columnHidden = false;
        await context.sync();
      } catch (err) {
        console.log("Can't show col due to: ", err);
      }
      await context.sync();
    }
    document.body.style.cursor = "default";
  });
}

async function Check_product_stage(productName, stagename) {
  console.log("productName in Check_product_stage: ", productName);
  console.log("stagename in Check_product_stage: ", stagename);
  await Excel.run(async (context) => {
    document.body.style.cursor = "wait";
    const masterSheet = context.workbook.worksheets.getItem("Masterfile");
    let usedRange;

    usedRange = masterSheet.getUsedRange();
    usedRange.load(["rowCount", "columnCount"]);
    await context.sync();
    const sheet = context.workbook.worksheets.getItem("Masterfile");
    const chunkSize = 1000;
    const totalRows = usedRange.rowCount;
    const totalCols = usedRange.columnCount;
    let allValues = [];
    for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
      const rowCount = Math.min(chunkSize, totalRows - startRow);
      const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
      range.load("values");
      await context.sync();
      allValues = allValues.concat(range.values);
    }
    let headers = allValues[0];
    console.log("headers: ", headers);

    //Insert Product as a Product Name
    let productColIndex = headers.indexOf(productName);
    console.log("productColINdex before add product name or stage: %d", productColIndex);
    logToConsole("productColINdex before add product name or stage: %d", productColIndex);
    let Allproduct_stage = [];
    let StartStageCol;
    let EndStageCol;
    let allstagescount;
    for (let i = 0; i <= headers.length; i++) {
      if (headers[i] === "Can remove (Y/N)") {
        StartStageCol = i;
      }
      if (headers[i] === "Lsl_typ") {
        EndStageCol = i;
      }
    }
    allstagescount = EndStageCol - StartStageCol - 1;
    let temp;
    if (allstagescount > 0) {
      for (let i = StartStageCol + 1; i < EndStageCol; i++) {
        const Procell = headers[i];
        const stageCell = allValues[1][i];
        if (Procell && Procell.trim() !== "") {
          Allproduct_stage.push({
            name: Procell.trim(),
            stage: stageCell,
          });
          temp = Procell.trim();
        } else {
          Allproduct_stage.push({
            name: temp,
            stage: stageCell,
          });
        }
      }
    }

    // If there is no same product name then insert it in
    if (productColIndex === -1) {
      const sheet = context.workbook.worksheets.getItem("Masterfile");
      //check if product name start with  T or P if T then show F,A if P then show W
      //if()
      const columnToInsert = sheet.getRange("F:F");
      columnToInsert.insert(Excel.InsertShiftDirection.right);
      const product_name_head = sheet.getRange("F1:F1");
      product_name_head.values = [[productName]];
      let Canremove_index = headers.indexOf("Can remove (Y/N)");
      if (Canremove_index < 0) {
        logToConsole("Can't find Can remove col");
        return;
      }
      let Lsl_typ_index = headers.indexOf("Lsl_typ");
      if (Lsl_typ_index < 0) {
        logToConsole("Can't find Lsl_typ col");
        return;
      }
      let Product_count = 0;
      for (let i = Canremove_index; i < Lsl_typ_index; i++) {
        let cell = usedRange.getCell(0, i);
        if (!isNaN(cell) || cell !== "") {
          Product_count++;
        }
      }
      const colors = ["#C6EFCE", "#FFEB9C", "#FFC7CE", "#e6cdfa"];
      const color = colors[Product_count % 4];
      product_name_head.format.fill.color = color;
      //add stage
      const stage_name_head = sheet.getRange("F2:F2");
      stage_name_head.values = [[stagename]];
      await context.sync();
    } else {
      //if product name is same then check if the stage is same
      const sheet = context.workbook.worksheets.getItem("Masterfile");
      await context.sync();
      const startCol = productColIndex;
      const stage_count = Allproduct_stage.filter((item) => item.name === productName).length; //how many stages does this product have
      console.log("product :  %s , stage count : %d", productName, stage_count);
      logToConsole("product : %s , stage count : %d", productName, stage_count);
      let columnToInsert = sheet.getRangeByIndexes(0, startCol + 1, 1, 1);
      columnToInsert.insert(Excel.InsertShiftDirection.right);
      usedRange = sheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();
      const stagerow = usedRange.rowCount;
      let stage_name_head = sheet.getRangeByIndexes(1, startCol + stage_count, stagerow, 1);
      stage_name_head.insert(Excel.InsertShiftDirection.right);
      usedRange = sheet.getUsedRange();
      await context.sync();
      const stageCell = sheet.getCell(1, startCol + stage_count);
      stageCell.values = [[stagename]];
      console.log("startcol for merge: %d stagecount for merge: %d", startCol, stage_count);
      logToConsole("startcol for merge: %d stagecount for merge: %d", startCol, stage_count);
      const range = sheet.getRangeByIndexes(0, startCol, 1, stage_count);
      range.values = Array(1).fill(Array(stage_count).fill(productName)); // ใส่ค่า productName ลงในทุกเซลล์ของ range Array(stage_count).fill(productName) สร้างอาร์เรย์ย่อยที่มี productName ซ้ำกัน stage_count ครั้ง เช่น ["ABC", "ABC", "ABC"] Array(1).fill(...) ทำให้กลายเป็น array 2 มิติ (1 แถว n คอลัมน์) → ซึ่งตรงกับ .values ที่ต้องการ
      range.merge();
      await context.sync();
    }
    document.body.style.cursor = "default";
  });
}
// write data of new tests
async function WriteNewtest(data) {
  console.log("data from EY oneproduct: ", data);
  await Excel.run(async (context) => {
    document.body.style.cursor = "wait";
    const sheet = context.workbook.worksheets.getItem("Masterfile");
    let usedRange = sheet.getUsedRange();
    usedRange.load(["rowCount", "columnCount"]);
    await context.sync();
    let chunkSize = 1000;
    let totalRows = usedRange.rowCount;
    let totalCols = usedRange.columnCount;
    let allValues = [];
    for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
      const rowCount = Math.min(chunkSize, totalRows - startRow);
      const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
      range.load("values");
      await context.sync();
      allValues = allValues.concat(range.values);
    }
    let headers = allValues[0];
    let TestnumColIndex = headers.indexOf("Test number");
    const SuiteColIndex = headers.indexOf("Suite name");
    const TestColIndex = headers.indexOf("Test name");
    if (TestnumColIndex === -1) {
      console.error("Can't find Test number column");
      logToConsole("Can't find Test number column");
      return;
    }
    if (SuiteColIndex === -1) {
      console.error("Can't find Suite name column");
      logToConsole("Can't find Suite name column");
      return;
    }
    if (TestColIndex === -1) {
      console.error("Can't find Test name column");
      logToConsole("Can't find Test name column");
      return;
    }
    const testNameRange = sheet.getRangeByIndexes(2, TestColIndex, allValues.length - 2, 1);
    testNameRange.load("values");
    await context.sync();
    logToConsole("Determined Allcolindex and testNamerange");
    let existingTestNames = [];
    try {
      existingTestNames = testNameRange.values.flat().filter((v) => v !== "");
    } catch (err) {
      console.error("Error while processing testNameRange.values:", err);
      return;
    }
    if (!Array.isArray(data)) {
      console.error("data isn't an array or isn't loaded");
      return;
    }
    let newTests = [];
    try {
      newTests = data.filter((item) => !existingTestNames.includes(item.test_name));
      console.log("newTests from EY: ", newTests);
    } catch (err) {
      console.error("Error while data.filter", err);
      return;
    }
    if (!Array.isArray(allValues)) {
      console.error("allValues isn't an array");
      return;
    }
    let startRow = allValues.length;
    let suiteRange, testRange;
    let suiteValues = [];
    let testValues = [];
    if (newTests.length > 0) {
      const testNumbers = newTests.map((t) => [t?.test_number ?? ""]);
      logToConsole("newTests.length = %d", newTests.length);
      if (TestnumColIndex === -1) {
        logToConsole("Can't find Test number column");
        return;
      }
      const writeRange = sheet.getRangeByIndexes(startRow, TestnumColIndex, newTests.length, 1);
      writeRange.values = testNumbers;
      await context.sync();
      suiteRange = sheet.getRangeByIndexes(startRow, SuiteColIndex, newTests.length, 1);
      testRange = sheet.getRangeByIndexes(startRow, TestColIndex, newTests.length, 1);
      suiteValues = newTests.map((t) => [t.suite_name]);
      testValues = newTests.map((t) => [t.test_name]);
      suiteRange.values = suiteValues;
      testRange.values = testValues;
      await context.sync();
    } else {
      logToConsole("There's no new tests");
    }
    document.body.style.cursor = "default";
  });
}
//Process Y/N values of data from EY
async function YN(data, productName, stagename) {
  await Excel.run(async (context) => {
    document.body.style.cursor = "wait";
    const sheet = context.workbook.worksheets.getItem("Masterfile");
    let usedRange = sheet.getUsedRange();
    usedRange.load(["rowCount", "columnCount"]);
    await context.sync();

    let chunkSize = 1000;
    let totalRows = usedRange.rowCount;
    let totalCols = usedRange.columnCount;
    let allValues = [];
    for (let startRow = 0; startRow < totalRows; startRow += chunkSize) {
      const rowCount = Math.min(chunkSize, totalRows - startRow);
      const range = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
      range.load("values");
      await context.sync();
      allValues = allValues.concat(range.values);
    }
    let headers = allValues[0];
    let Allproduct_stage = [];
    let StartStageCol;
    let EndStageCol;
    let allstagescount;
    for (let i = 0; i <= headers.length; i++) {
      if (headers[i] === "Can remove (Y/N)") {
        StartStageCol = i;
      }
      if (headers[i] === "Lsl_typ") {
        EndStageCol = i;
      }
    }
    allstagescount = EndStageCol - StartStageCol - 1;
    let temp;
    if (allstagescount > 0) {
      for (let i = StartStageCol + 1; i < EndStageCol; i++) {
        const Procell = headers[i];
        const stageCell = allValues[1][i];
        if (Procell && Procell.trim() !== "") {
          Allproduct_stage.push({
            name: Procell.trim(),
            stage: stageCell,
          });
          temp = Procell.trim();
        } else {
          Allproduct_stage.push({
            name: temp,
            stage: stageCell,
          });
        }
      }
    }
    let TestnumColIndex = headers.indexOf("Test number");
    const SuiteColIndex = headers.indexOf("Suite name");
    const TestColIndex = headers.indexOf("Test name");
    if (TestnumColIndex === -1) {
      console.error("Test number column is not found");
      logToConsole("Test number column is not found");
      return;
    }
    if (SuiteColIndex === -1) {
      console.error("Suite name column is not found");
      logToConsole("Suite name column is not found");
      return;
    }
    if (TestColIndex === -1) {
      console.error("Test name column is not found");
      logToConsole("Test name column is not found");
      return;
    }
    let productColIndex = headers.indexOf(productName);
    if (productColIndex === -1) {
      console.error("Can't find:", productName);
      logToConsole("Can't find:", productName);
      return;
    }
    let stage_count = Allproduct_stage.filter((item) => item.name === productName).length;
    let stage_array_index;
    let stage_range = sheet.getRangeByIndexes(1, productColIndex, 1, stage_count);
    stage_range.load("values");
    await context.sync();
    for (let i = 0; i <= stage_count; i++) {
      console.log("stage %d = %s", i, stage_range.values[0][i]);
      if (stage_range.values[0][i] === stagename) {
        stage_array_index = i;
        break;
      }
    }
    if (stage_array_index === undefined) {
      console.error("Can't find:", stagename);
      logToConsole("Can't find:", stagename);
    }
    console.log(
      "productColIndex: %d, stage_count: %d, stageArrayIndex: %d ",
      productColIndex,
      stage_count,
      stage_array_index
    );
    logToConsole(
      "productColIndex: %d, stage_count: %d, stageArrayIndex: %d ",
      productColIndex,
      stage_count,
      stage_array_index
    );

    const testNameRangeAll = sheet.getRangeByIndexes(2, TestColIndex, allValues.length - 2, 1);
    testNameRangeAll.load("values");
    await context.sync();
    // create YNValues by mapping test_name -> YN_check
    const allTestNames = testNameRangeAll.values.map((row) => row[0]);
    logToConsole("allTestNames length : %d", allTestNames.length);
    let YNValues = [];
    try {
      YNValues = allTestNames.map((testName) => {
        const match = data.find((item) => item.test_name === testName);
        return [match ? match.YN_check : ""];
      });
    } catch (err) {
      console.error("Error while processing YNValues:", err);
      logToConsole("Error while processing YNValues: %s", err.message || err);
    }
    logToConsole("YNcolIndex : %d", productColIndex + stage_array_index);
    let YNRange = sheet.getRangeByIndexes(
      2,
      productColIndex + stage_array_index,
      YNValues.length,
      1
    );
    YNRange.load("values");
    await context.sync();

    if (YNValues.length === 0) {
      console.warn("There's no Y/N check data");
      logToConsole("There's no Y/N check data");
    } else {
      console.log("YN.length of %s %s is %d", productName, stagename, YNValues.length);
      logToConsole("YN.length of %s %s is %d", productName, stagename, YNValues.length);
    }
    YNRange.values = YNValues;
    await context.sync();
    // loop for add green color and add N for null cell (not yet)
    const IsUsedIndex = headers.indexOf("Is used (Y/N)");
    let IsUsedDataRange = sheet.getRangeByIndexes(2, IsUsedIndex, YNRange.values.length, 1);
    IsUsedDataRange.load("values");
    await context.sync();
    let IsUsedData = IsUsedDataRange.values;
    // If IsUsedData is null then create new empty array to prevent undefine error
    if (!Array.isArray(IsUsedData) || IsUsedData.length === 0) {
      IsUsedData = Array.from({ length: YNRange.values.length }, () => [""]);
    }
    for (let i = 0; i < YNRange.values.length; i++) {
      if (YNRange.values[i][0] === "Y") {
        if (IsUsedData[i][0] === "Partial" || IsUsedData[i][0] === "No") {
          IsUsedData[i][0] = "Partial";
        } else if (IsUsedData[i][0] === "") {
          IsUsedData[i][0] = "All";
        }
      } else {
        if (IsUsedData[i][0] === "All" || IsUsedData[i][0] === "Partial") {
          IsUsedData[i][0] = "Partial";
        } else IsUsedData[i][0] = "No";
      }
    }
    IsUsedDataRange.values = IsUsedData;
    await context.sync();
    //conditional formatting color
    let conditionalFormat = YNRange.conditionalFormats.add(
      Excel.ConditionalFormatType.containsText
    );
    conditionalFormat.textComparison.format.fill.color = "#C6EFCE";
    conditionalFormat.textComparison.rule = {
      operator: Excel.ConditionalTextOperator.contains,
      text: "Y",
    };
    IsUsedDataRange.conditionalFormats.load("count");
    await context.sync();

    for (let i = IsUsedDataRange.conditionalFormats.count - 1; i >= 0; i--) {
      IsUsedDataRange.conditionalFormats.getItemAt(i).delete();
    }
    await context.sync();
    const IsUsedkeywords = ["Partial", "All"];
    const colors = ["#FFEB9C", "#C6EFCE"];

    for (let i = 0; i < IsUsedkeywords.length; i++) {
      const word = IsUsedkeywords[i];
      const color = colors[i];

      const conditionalFormat = IsUsedDataRange.conditionalFormats.add(
        Excel.ConditionalFormatType.containsText
      );
      conditionalFormat.textComparison.format.fill.color = color;
      conditionalFormat.textComparison.rule = {
        operator: Excel.ConditionalTextOperator.contains,
        text: word,
      };
    }

    await context.sync();
    document.body.style.cursor = "default";
  });
}
