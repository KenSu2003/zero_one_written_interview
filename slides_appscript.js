// Global variable for API key
let API_KEY = "AIzaSyDJjsuMtEBXPU2zwtTkMlzoEA7xPf5nqJA";

function createGoogleSlidesPresentation() {
  // Get the spreadsheet with keyword data
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keywordSheet = ss.getSheetByName("工作表1") || ss.getSheets()[0]; // Try to find the sheet with keyword data
  
  // Create a new presentation
  const presentation = SlidesApp.create("滴雞精關鍵字分析報告");
  const presentationId = presentation.getId();
  
  // Get the first slide (title slide)
  const titleSlide = presentation.getSlides()[0];
  titleSlide.getShapes()[0].getText().setText("滴雞精關鍵字分析報告");
  titleSlide.getShapes()[1].getText().setText("基於零一筆試_關鍵字模擬數據");
  
  // Get all data including the checkbox column (G)
  const dataRange = keywordSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  
  // Ensure column G has checkboxes (if not already set up)
  ensureCheckboxesInColumnG(keywordSheet, data.length);
  
  // Create a slide for selected top keywords
  const selectedRows = [];
  
  // Find rows with checked checkboxes (column G)
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === true) { // Column G (index 6) has checkbox
      selectedRows.push(i);
    }
  }
  
  if (selectedRows.length > 0) {
    const topKeywordsSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS);
    topKeywordsSlide.getShapes()[0].getText().setText("已選關鍵字");
    
    // Create a table for the selected data
    const leftColumn = topKeywordsSlide.getShapes()[1].getText();
    let tableContent = "";
    
    selectedRows.forEach((rowIndex, index) => {
      // Get data from the selected row (assuming columns B-F contain keyword data)
      const rowData = data[rowIndex].slice(1, 6); // B-F columns (indices 1-5)
      const [keyword, impressions, clicks, ctr, position] = rowData;
      
      tableContent += `${index + 1}. ${keyword}\n`;
      tableContent += `   搜尋量: ${impressions}, 點擊: ${clicks}\n`;
      tableContent += `   CTR: ${ctr}, 排名: ${position}\n\n`;
    });
    
    leftColumn.setText(tableContent);
    
    // Create individual slides for each selected row
    selectedRows.forEach(rowIndex => {
      const rowData = data[rowIndex].slice(1, 6); // B-F columns (indices 1-5)
      const [keyword, impressions, clicks, ctr, position] = rowData;
      
      const detailSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      detailSlide.getShapes()[0].getText().setText(`關鍵字詳情: ${keyword}`);
      
      const bodyText = detailSlide.getShapes()[1].getText();
      bodyText.setText(
        `搜尋量: ${impressions}\n` +
        `點擊數: ${clicks}\n` +
        `點擊率: ${ctr}\n` +
        `平均排名: ${position}\n\n` +
        `分析建議: 此關鍵字可用於...`
      );
    });
  } else {
    // If no rows are selected, create a slide indicating this
    const noSelectionSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    noSelectionSlide.getShapes()[0].getText().setText("尚未選擇關鍵字");
    noSelectionSlide.getShapes()[1].getText().setText("請在試算表中的 G 欄勾選要包含在簡報中的關鍵字。");
  }
  
  // Create a slide for Gemini API analysis if any rows are selected
  if (selectedRows.length > 0) {
    const geminiSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    geminiSlide.getShapes()[0].getText().setText("Gemini AI 關鍵字分析");
    
    // Extract data for selected rows only
    const selectedKeywordData = selectedRows.map(rowIndex => data[rowIndex].slice(1, 6));
    
    // Call Gemini API with only the selected data
    const geminiAnalysis = getGeminiAnalysis(selectedKeywordData);
    geminiSlide.getShapes()[1].getText().setText(geminiAnalysis);
  }
  
  // Return the presentation URL
  return `https://docs.google.com/presentation/d/${presentationId}/edit`;
}

function getGeminiAnalysis(keywordData) {
  // Get content from the "ZeroOne 關鍵字數據" sheet, cell D4
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let contentForAnalysis = "";
  
  try {
    const zeroOneSheet = ss.getSheetByName("ZeroOne 關鍵字數據");
    if (zeroOneSheet) {
      contentForAnalysis = zeroOneSheet.getRange("D4").getValue();
      // If content is empty, add a note about it
      if (!contentForAnalysis || contentForAnalysis.trim() === "") {
        contentForAnalysis = "注意：未提供內容分析文本。";
      }
    } else {
      contentForAnalysis = "注意：找不到「ZeroOne 關鍵字數據」工作表。";
    }
  } catch (e) {
    Logger.log("Error accessing ZeroOne 關鍵字數據 sheet: " + e.toString());
    contentForAnalysis = "注意：讀取內容時發生錯誤。";
  }
  
  // Build a prompt for the Gemini API similar to the Python script
  let prompt = "以下是零一筆試的關鍵字模擬數據，請你：\n";
  // Add the content for analysis
  prompt += "\nPlease read through this content:\n" + contentForAnalysis;
  prompt += "\nThe goal is to boost the content's SEO and make it more engaging for the readers as well as increase impressions and clicks.\n";
  prompt += "\nGiven the content, use the keywords that are checked on column G to decide which section the keyword it should be placed in.\n";
  prompt += "\nAfter making such identification, proceed to make recommendation on how to use the keword or add a paragraph in that section.\n";
  
  // Format data similar to the Python script
  // Check if keywordData exists and is an array
  if (keywordData && Array.isArray(keywordData) && keywordData.length > 0) {
    prompt += "\nSelected keywords:\n";
    keywordData.forEach((row, index) => {
      if (index < 10) { // Only use top 10 rows like Python script
        const [keyword, impressions, clicks, ctr, position] = row;
        prompt += `關鍵字：${keyword} | 點擊率：${ctr} | 搜尋量：${impressions} | 平均排名：${position}\n`;
      }
    });
  } else {
    prompt += "\nNo keywords selected. Please check at least one keyword in column G.\n";
  }

  
  if (!API_KEY) {
    return "請先設置 Gemini API Key。點擊選單中的「設置 API Key」選項。";
  }
  
  // Using gemini-1.5-pro which is available in the REST API
  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-1.5-pro:generateContent?key=${API_KEY}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 1024
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseJson = JSON.parse(response.getContentText());
    
    console.log("API Response:", JSON.stringify(responseJson)); // Debug logging
    
    if (responseJson.candidates && responseJson.candidates.length > 0 && 
        responseJson.candidates[0].content && responseJson.candidates[0].content.parts) {
      return responseJson.candidates[0].content.parts[0].text;
    } else {
      // Log the full response for debugging
      Logger.log("Unexpected response structure: " + JSON.stringify(responseJson));
      return "Error: Unexpected response structure from Gemini API.";
    }
  } catch (e) {
    Logger.log("Error calling Gemini API: " + e.toString());
    return "Error calling Gemini API: " + e.toString();
  }
}

// Function to set up the Gemini API key
function setupGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Gemini API 設定',
    '請輸入您的 Gemini API Key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
      ui.alert('成功', 'Gemini API Key 已成功設定', ui.ButtonSet.OK);
      return true;
    } else {
      ui.alert('錯誤', '請輸入有效的 API Key', ui.ButtonSet.OK);
      return false;
    }
  }
  return false;
}

// Function to ensure column G has checkboxes
function ensureCheckboxesInColumnG(sheet, rowCount) {
  // Set header for column G if not already set
  const gHeader = sheet.getRange(1, 7).getValue();
  if (!gHeader) {
    sheet.getRange(1, 7).setValue("包含在簡報");
  }
  
  // Add checkboxes to all data rows
  if (rowCount > 1) {
    const checkboxRange = sheet.getRange(2, 7, rowCount - 1, 1);
    
    // Check if data validation already exists
    const existingValidation = checkboxRange.getDataValidation();
    if (!existingValidation) {
      // Create checkbox data validation rule
      const rule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      
      checkboxRange.setDataValidation(rule);
    }
  }
}

// Add a function to create an onEdit trigger to auto-update presentation
function createOnEditTrigger() {
  // Delete any existing triggers first to avoid duplicates
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActiveSpreadsheet());
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onEditCheckbox") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger
  ScriptApp.newTrigger("onEditCheckbox")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  SpreadsheetApp.getUi().alert("自動更新觸發器已設置。當勾選方塊改變時，簡報將自動更新。");
}

// Function that runs when a checkbox is edited
function onEditCheckbox(e) {
  // Check if the edit was in column G (checkboxes)
  if (e.range.getColumn() === 7 && e.range.getRow() > 1) {
    // Get the active spreadsheet
    const ss = e.source;
    const sheet = e.range.getSheet();
    
    // Only proceed if we're in the correct sheet
    if (sheet.getName() === "工作表1" || sheet.getIndex() === 1) {
      // Show a toast notification
      ss.toast("正在更新簡報...", "自動更新", 3);
      
      // Create or update the presentation
      const presentationUrl = createGoogleSlidesPresentation();
      
      // Show a toast with the link
      ss.toast(`簡報已更新: ${presentationUrl}`, "完成", 5);
    }
  }
}