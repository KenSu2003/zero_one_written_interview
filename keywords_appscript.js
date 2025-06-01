// Global variable for API key
let API_KEY = "AIzaSyDJjsuMtEBXPU2zwtTkMlzoEA7xPf5nqJA";
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
  
    // Create a map to store results grouped by keyword
    const keywordMap = {};
    keywordList.forEach(kw => {
      keywordMap[kw] = [];
    });
    
    // Start from row 1 (skipping header)
    for (let i = 1; i < data.length; i++) {
      const title = data[i][titleIndex];
      const content = data[i][contentIndex];
      
      if (!content || content.trim().length === 0) continue;
  
      // Split by periods (。), exclamation marks (！), question marks (？), and newlines
      const paragraphs = content.split(/[。！？\n]+/);
  
      let paraCount = 1;
      for (let para of paragraphs) {
        para = para.trim();
        if (!para || para.length < 5) continue; // Skip very short segments
        
        keywordList.forEach(kw => {
          // Create a regex that matches the whole keyword
          const regex = new RegExp(kw, "g");
          const matches = para.match(regex);
          
          if (matches && matches.length > 0) {
            // Add each instance to the keyword map
            keywordMap[kw].push({
              title: title,
              paragraph: para,
              paraNum: paraCount,
              count: matches.length
            });
          }
        });
        
        paraCount++;
      }
    }
  
    // Prepare output data
    const output = [["關鍵字", "文章標題", "段落編號", "出現次數", "段落內容"]];
    
    for (const keyword in keywordMap) {
      const instances = keywordMap[keyword];
      
      // Skip keywords with no matches
      if (instances.length === 0) continue;
      
      // Add each instance as a row
      instances.forEach(instance => {
        output.push([
          keyword,
          instance.title,
          instance.paraNum,
          instance.count,
          instance.paragraph
        ]);
      });
    }
  
    // Output to the sheet
    outputSheet.clear();
    if (output.length > 1) {
      outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
      
      // Format header row
      outputSheet.getRange(1, 1, 1, output[0].length).setFontWeight("bold");
      
      // Auto-size columns for better readability
      outputSheet.autoResizeColumns(1, output[0].length);
      
      // Set text wrapping for paragraph content
      outputSheet.getRange(2, 5, output.length - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    } else {
      outputSheet.getRange(1, 1).setValue("No matching content found");
    }
    
    Logger.log("Analysis complete. Found " + (output.length - 1) + " keyword instances.");
  }