// Global variable for API key
let API_KEY = "AIzaSyDJjsuMtEBXPU2zwtTkMlzoEA7xPf5nqJA";

function createGoogleSlidesPresentation() {
  try {
    // Get the spreadsheet with keyword data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const keywordSheet = ss.getSheetByName("工作表1") || ss.getSheets()[0]; // Try to find the sheet with keyword data
    
    // Check if we already have a presentation ID stored
    let presentation;
    let presentationId = PropertiesService.getScriptProperties().getProperty('SLIDES_PRESENTATION_ID');
    Logger.log("Retrieved presentation ID from properties: " + presentationId);
    
    let isNewPresentation = false;
    
    try {
      // Try to open existing presentation if we have an ID
      if (presentationId) {
        try {
          Logger.log("Attempting to open existing presentation with ID: " + presentationId);
          presentation = SlidesApp.openById(presentationId);
          Logger.log("Successfully opened existing presentation");
          
          // Clear all slides except the first one (title slide)
          const slides = presentation.getSlides();
          Logger.log("Found " + slides.length + " slides in the presentation");
          
          for (let i = slides.length - 1; i > 0; i--) {
            slides[i].remove();
          }
          Logger.log("Removed all slides except title slide");
          
          // Update the title slide - make sure to properly access the title and subtitle placeholders
          const titleSlide = slides[0];
          
          // Get all shapes on the title slide
          const shapes = titleSlide.getShapes();
          Logger.log("Found " + shapes.length + " shapes on title slide");
          
          // Find and update the title placeholder
          let titleUpdated = false;
          let subtitleUpdated = false;
          
          for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];
            const placeholder = shape.getPlaceholderType();
            
            if (placeholder === SlidesApp.PlaceholderType.TITLE) {
              shape.getText().setText("滴雞精關鍵字分析報告");
              titleUpdated = true;
              Logger.log("Updated title placeholder");
            } else if (placeholder === SlidesApp.PlaceholderType.SUBTITLE) {
              shape.getText().setText("基於零一筆試_關鍵字模擬數據");
              subtitleUpdated = true;
              Logger.log("Updated subtitle placeholder");
            }
          }
          
          // If we couldn't find the placeholders, try the first two shapes as a fallback
          if (!titleUpdated && shapes.length > 0) {
            shapes[0].getText().setText("滴雞精關鍵字分析報告");
            Logger.log("Updated title using first shape");
          }
          
          if (!subtitleUpdated && shapes.length > 1) {
            shapes[1].getText().setText("基於零一筆試_關鍵字模擬數據");
            Logger.log("Updated subtitle using second shape");
          }
        } catch (e) {
          // If we can't open the presentation (e.g., it was deleted), create a new one
          Logger.log("Could not open existing presentation: " + e.toString());
          presentation = SlidesApp.create("滴雞精關鍵字分析報告");
          presentationId = presentation.getId();
          PropertiesService.getScriptProperties().setProperty('SLIDES_PRESENTATION_ID', presentationId);
          Logger.log("Created new presentation with ID: " + presentationId);
          isNewPresentation = true;
        }
      } else {
        // Create a new presentation if we don't have an ID stored
        Logger.log("No presentation ID found, creating new presentation");
        presentation = SlidesApp.create("滴雞精關鍵字分析報告");
        presentationId = presentation.getId();
        PropertiesService.getScriptProperties().setProperty('SLIDES_PRESENTATION_ID', presentationId);
        Logger.log("Created new presentation with ID: " + presentationId);
        isNewPresentation = true;
      }
      
      // For new presentations, ensure the title slide is properly set up
      if (isNewPresentation) {
        Logger.log("Setting up title slide for new presentation");
        
        // Force a small delay to ensure the presentation is fully created
        Utilities.sleep(1000);
        
        // Reload the presentation to ensure we have the latest version
        presentation = SlidesApp.openById(presentationId);
        
        // Get the title slide
        const titleSlide = presentation.getSlides()[0];
        const shapes = titleSlide.getShapes();
        
        Logger.log("New presentation has " + shapes.length + " shapes on title slide");
        
        let titleUpdated = false;
        let subtitleUpdated = false;
        
        // Try to find and update placeholders
        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          try {
            const placeholder = shape.getPlaceholderType();
            
            if (placeholder === SlidesApp.PlaceholderType.TITLE) {
              shape.getText().setText("滴雞精關鍵字分析報告");
              titleUpdated = true;
              Logger.log("Set title on new presentation using placeholder");
            } else if (placeholder === SlidesApp.PlaceholderType.SUBTITLE) {
              shape.getText().setText("基於零一筆試_關鍵字模擬數據");
              subtitleUpdated = true;
              Logger.log("Set subtitle on new presentation using placeholder");
            }
          } catch (e) {
            Logger.log("Error checking placeholder for shape " + i + ": " + e.toString());
          }
        }
        
        // If placeholders not found, use index-based approach
        if (!titleUpdated && shapes.length > 0) {
          shapes[0].getText().setText("滴雞精關鍵字分析報告");
          Logger.log("Set title on new presentation using first shape");
        }
        
        if (!subtitleUpdated && shapes.length > 1) {
          shapes[1].getText().setText("基於零一筆試_關鍵字模擬數據");
          Logger.log("Set subtitle on new presentation using second shape");
        }
        
        // If still not updated, try creating a new title and subtitle
        if ((!titleUpdated || !subtitleUpdated) && shapes.length === 0) {
          Logger.log("No shapes found on title slide, creating new title and subtitle");
          
          // Create a title text box
          const titleBox = titleSlide.insertTextBox("滴雞精關鍵字分析報告", 100, 100, 400, 50);
          titleBox.getText().getTextStyle().setBold(true).setFontSize(24);
          
          // Create a subtitle text box
          const subtitleBox = titleSlide.insertTextBox("基於零一筆試_關鍵字模擬數據", 100, 160, 400, 30);
          subtitleBox.getText().getTextStyle().setFontSize(18);
        }
      }
    } catch (e) {
      // If there's any error in the process, create a new presentation
      Logger.log("Error handling presentation: " + e.toString());
      presentation = SlidesApp.create("滴雞精關鍵字分析報告");
      presentationId = presentation.getId();
      PropertiesService.getScriptProperties().setProperty('SLIDES_PRESENTATION_ID', presentationId);
      Logger.log("Created new presentation with ID: " + presentationId);
      
      // Set up title slide for the new presentation
      Utilities.sleep(1000);
      presentation = SlidesApp.openById(presentationId);
      const titleSlide = presentation.getSlides()[0];
      const shapes = titleSlide.getShapes();
      
      if (shapes.length > 0) {
        shapes[0].getText().setText("滴雞精關鍵字分析報告");
        Logger.log("Set title on error recovery presentation");
      }
      
      if (shapes.length > 1) {
        shapes[1].getText().setText("基於零一筆試_關鍵字模擬數據");
        Logger.log("Set subtitle on error recovery presentation");
      }
    }
    
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
    Logger.log("Found " + selectedRows.length + " selected rows");
    
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
      Logger.log("Created top keywords slide");
      
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
      Logger.log("Created " + selectedRows.length + " detail slides");
    } else {
      // If no rows are selected, create a slide indicating this
      const noSelectionSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      noSelectionSlide.getShapes()[0].getText().setText("尚未選擇關鍵字");
      noSelectionSlide.getShapes()[1].getText().setText("請在試算表中的 G 欄勾選要包含在簡報中的關鍵字。");
      Logger.log("Created no selection slide");
    }
    
    // Create a slide for Gemini API analysis if any rows are selected
    if (selectedRows.length > 0) {
      try {
        const geminiSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
        geminiSlide.getShapes()[0].getText().setText("Gemini AI 關鍵字分析");
        
        // Extract data for selected rows only
        const selectedKeywordData = selectedRows.map(rowIndex => data[rowIndex].slice(1, 6));
        
        Logger.log("Calling Gemini API for analysis");
        // Call Gemini API with only the selected data
        const geminiAnalysis = getGeminiAnalysis(selectedKeywordData);
        geminiSlide.getShapes()[1].getText().setText(geminiAnalysis);
        Logger.log("Added Gemini analysis slide");
      } catch (e) {
        Logger.log("Error creating Gemini slide: " + e.toString());
        // Continue even if Gemini API fails
      }
    }
    
    // Return the presentation URL
    Logger.log("Returning presentation URL: " + `https://docs.google.com/presentation/d/${presentationId}/edit`);
    return `https://docs.google.com/presentation/d/${presentationId}/edit`;
  } catch (e) {
    Logger.log("Critical error in createGoogleSlidesPresentation: " + e.toString());
    throw e; // Re-throw to show error to user
  }
}

function getGeminiAnalysis(keywordData) {
  // Get content from the external file "ZeroOne 關鍵字數據", cell D3 of Sheet1
  let contentForAnalysis = "";
  
  try {
    // Access the external file by name
    const files = DriveApp.getFilesByName("ZeroOne 關鍵字數據");
    
    if (files.hasNext()) {
      const file = files.next();
      const externalSS = SpreadsheetApp.open(file);
      const sheet1 = externalSS.getSheetByName("Sheet1") || externalSS.getSheets()[0];
      
      if (sheet1) {
        contentForAnalysis = sheet1.getRange("D3").getValue();
        Logger.log("Successfully retrieved content from external file: " + contentForAnalysis.substring(0, 50) + "...");
      } else {
        contentForAnalysis = "Error: Could not find Sheet1 in the external file";
        Logger.log("Could not find Sheet1 in the external file");
      }
    } else {
      contentForAnalysis = "Error: Could not find the external file ZeroOne 關鍵字數據";
      Logger.log("Could not find the external file ZeroOne 關鍵字數據");
    }
    
    // If content is empty, add a note about it
    if (!contentForAnalysis || contentForAnalysis.toString().trim() === "") {
      contentForAnalysis = "注意：外部檔案中未提供內容分析文本。";
    }
  } catch (e) {
    Logger.log("Error accessing external file: " + e.toString());
    contentForAnalysis = "注意：讀取外部檔案時發生錯誤: " + e.toString();
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

// Add a function to create a new presentation (if needed)
function createNewPresentation() {
  // Create a new presentation and store its ID
  const presentation = SlidesApp.create("滴雞精關鍵字分析報告");
  const presentationId = presentation.getId();
  PropertiesService.getScriptProperties().setProperty('SLIDES_PRESENTATION_ID', presentationId);
  
  SpreadsheetApp.getUi().alert("已創建新的簡報。下次更新時將使用此簡報。");
  
  return `https://docs.google.com/presentation/d/${presentationId}/edit`;
}