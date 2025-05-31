function analyzeAllArticles() {
  const keywordList = ["滴雞精", "營養", "推薦", "補身", "品牌", "熬煮", "口感", "純正", "健康", "養生", "cc"];
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1"); // adjust if needed
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("分析結果") ||
                      SpreadsheetApp.getActiveSpreadsheet().insertSheet("分析結果");

  // Get all data including headers
  const data = inputSheet.getDataRange().getValues();
  
  // Find column indexes based on headers
  const headers = data[0];
  const titleIndex = headers.indexOf("Title");
  const contentIndex = headers.indexOf("Content");
  
  if (titleIndex === -1 || contentIndex === -1) {
    Logger.log("Required columns 'Title' or 'Content' not found");
    return;
  }

  // Create a map to store results grouped by title
  const articleMap = {};
  
  // Start from row 1 (skipping header)
  for (let i = 1; i < data.length; i++) {
    const title = data[i][titleIndex];
    const content = data[i][contentIndex];
    
    if (!content || content.trim().length === 0) continue;

    // Initialize article entry if it doesn't exist
    if (!articleMap[title]) {
      articleMap[title] = {
        paragraphData: [],
        keywordCounts: {}
      };
    }

    // Split by periods (。), exclamation marks (！), question marks (？), and newlines
    const paragraphs = content.split(/[。！？\n]+/);

    let paraCount = 1;
    for (let para of paragraphs) {
      para = para.trim();
      if (!para || para.length < 5) continue; // Skip very short segments
      
      let matches = [];
      let hasKeyword = false;
      
      keywordList.forEach(kw => {
        // Create a regex that matches the whole keyword
        const regex = new RegExp(kw, "g");
        const count = (para.match(regex) || []).length;
        
        if (count > 0) {
          matches.push(`${kw}(${count})`);
          hasKeyword = true;
          
          // Update total keyword count for this article
          articleMap[title].keywordCounts[kw] = (articleMap[title].keywordCounts[kw] || 0) + count;
        }
      });

      if (hasKeyword) {
        articleMap[title].paragraphData.push({
          paraNum: paraCount,
          content: para,
          keywords: matches.join(", ")
        });
      }
      paraCount++;
    }
  }

  // Prepare output data
  const output = [["文章標題", "關鍵字出現次數", "段落詳情"]];
  
  for (const title in articleMap) {
    const article = articleMap[title];
    
    // Skip articles with no keyword matches
    if (article.paragraphData.length === 0) continue;
    
    // Format keyword counts
    const keywordSummary = Object.entries(article.keywordCounts)
      .map(([kw, count]) => `${kw}(${count})`)
      .join(", ");
    
    // Format paragraph details
    const paragraphDetails = article.paragraphData
      .map(p => `段落${p.paraNum}: ${p.content} [${p.keywords}]`)
      .join("\n\n");
    
    output.push([title, keywordSummary, paragraphDetails]);
  }

  // Output to the sheet
  outputSheet.clear();
  if (output.length > 1) {
    outputSheet.getRange(1, 1, output.length, 3).setValues(output);
    
    // Format header row
    outputSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    
    // Auto-size columns for better readability
    outputSheet.autoResizeColumns(1, 3);
    
    // Set text wrapping for paragraph details
    outputSheet.getRange(2, 3, output.length - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  } else {
    outputSheet.getRange(1, 1).setValue("No matching content found");
  }
  
  Logger.log("Analysis complete. Found " + (output.length - 1) + " articles containing keywords.");
}

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
  
  // Create a slide for rows 2-10 (top keywords)
  const topKeywordsSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS);
  topKeywordsSlide.getShapes()[0].getText().setText("熱門關鍵字 (Top 9)");
  
  // Get data from rows 2-10
  const keywordData = keywordSheet.getRange(2, 2, 9, 5).getValues(); // B2:F10
  
  // Create a table for the data
  const leftColumn = topKeywordsSlide.getShapes()[1].getText();
  let tableContent = "";
  keywordData.forEach((row, index) => {
    const [keyword, impressions, clicks, ctr, position] = row;
    tableContent += `${index + 1}. ${keyword}\n`;
    tableContent += `   搜尋量: ${impressions}, 點擊: ${clicks}\n`;
    tableContent += `   CTR: ${ctr}, 排名: ${position}\n\n`;
  });
  leftColumn.setText(tableContent);
  
  // Create slides for specific rows (13, 16, 19)
  const specificRows = [13, 16, 19]; // These are rows 14, 17, 20 in the sheet (1-indexed)
  
  specificRows.forEach(rowIndex => {
    const rowData = keywordSheet.getRange(rowIndex + 1, 2, 1, 5).getValues()[0]; // +1 because sheet is 1-indexed
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
  
  // Create a slide for Gemini API analysis
  const geminiSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  geminiSlide.getShapes()[0].getText().setText("Gemini AI 關鍵字分析");
  
  // Call Gemini API (placeholder - this would be implemented differently in a real scenario)
  // In AppScript, we would typically call an external API or use pre-generated analysis
  const geminiAnalysis = getGeminiAnalysis(keywordData);
  geminiSlide.getShapes()[1].getText().setText(geminiAnalysis);
  
  // Return the presentation URL
  return `https://docs.google.com/presentation/d/${presentationId}/edit`;
}

function getGeminiAnalysis(keywordData) {
  // Build a prompt for the Gemini API similar to the Python script
  let prompt = "以下是零一筆試的關鍵字模擬數據，請你：\n";
  prompt += "- 分析哪些關鍵字行銷潛力高\n";
  prompt += "- 建議放入內容中的方式與使用場景\n";
  prompt += "- 按優先順序排序推薦處理的關鍵字\n\n";
  
  // Format data similar to the Python script
  keywordData.forEach((row, index) => {
    if (index < 10) { // Only use top 10 rows like Python script
      const [keyword, impressions, clicks, ctr, position] = row;
      prompt += `關鍵字：${keyword} | 點擊率：${ctr} | 搜尋量：${impressions} | 平均排名：${position}\n`;
    }
  });
  
  // Put your Gemini API key here
  const API_KEY = "YOUR_GEMINI_API_KEY"; // Replace with your actual API key
  
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