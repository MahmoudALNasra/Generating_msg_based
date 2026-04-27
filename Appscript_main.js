// CONFIGURATION - EASILY CHANGE THESE VARIABLES
const CONFIG = {
  arabicName: "محمود النصرة", // Change Arabic name here
  englishName: "Mahmoud Alnasra", // Change English name here
  companyName: "بوستجيتال", // Company name in Arabic
  companyNameEnglish: "boostigital", // Company name in English
  phoneNumber: "+96178738309", // Your phone number
  pagespeedApiKey: "" // Add your Google PageSpeed API key here
};

// GLOBAL CACHE AND QUOTA TRACKER (per execution)
const pageSpeedCache = {};
let urlFetchQuotaExhausted = false;
function checkQuotaUsage() {
  const url = "https://www.googleapis.com/discovery/v1/apis";
  try {
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    Logger.log("✅ URL Fetch quota still available.");
  } catch (e) {
    Logger.log("❌ URL Fetch quota likely exhausted: " + e.toString());
  }
}
function resetQuotaFlag() {
  clearQuotaExhausted();
  Logger.log('✅ Quota exhausted flag cleared. You can now use the new API key.');
}
function testNewKey() {
  const testUrl = "http://www.nahdionline.com/"; // Replace with a real URL
  const result = getConservativePageSpeedScore(testUrl);
  Logger.log(result ? `Score: ${result.score}` : "❌ Failed – check API key and enable PageSpeed API.");
}
// PERSISTENT QUOTA TRACKING (across script runs)
function isQuotaExhaustedToday() {
  const props = PropertiesService.getScriptProperties();
  const lastExhaustion = props.getProperty('PAGESPEED_QUOTA_EXHAUSTED_DATE');
  if (!lastExhaustion) return false;
  const today = new Date().toDateString();
  return lastExhaustion === today;
}

function setQuotaExhaustedToday() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('PAGESPEED_QUOTA_EXHAUSTED_DATE', new Date().toDateString());
}

function clearQuotaExhausted() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('PAGESPEED_QUOTA_EXHAUSTED_DATE');
}

// MAIN FUNCTION - RUN THIS FOR ALL SHEETS WITH RESUME CAPABILITY
function generateBilingualMessagesComplete() {
  const sheetNames = [
    // 'ELECTRONICS STORE_filtered',
    // 'PHARMACY_filtered', 
    // 'CAR DEALER_filtered',
    // 'REAL ESTATE AGENCY_filtered',
    // 'hotels without booking domain',
    // 'SUPERMARKET_filtered',
    // 'BEAUTY SALON_filtered',
    // 'CAFE_filtered',
    'GYM_filtered',
    // 'RESTAURANT_filtered'
  ];
  
  // Reset per-execution state
  Object.keys(pageSpeedCache).forEach(key => delete pageSpeedCache[key]);
  urlFetchQuotaExhausted = false;
  
  let totalMessages = 0;
  
  for (const sheetName of sheetNames) {
    Logger.log(`Generating complete bilingual messages for: ${sheetName}`);
    const sheetResult = generateMessagesForSheetWithResume(sheetName);
    totalMessages += sheetResult.generated;
    Logger.log(`Completed: ${sheetName} - ${sheetResult.generated} messages, ${sheetResult.skipped} skipped`);
  }
  
  Logger.log(`TOTAL: ${totalMessages} messages generated across all sheets`);
  return totalMessages;
}

// ENHANCED MESSAGE GENERATOR WITH PERSISTENT QUOTA HANDLING
// ENHANCED MESSAGE GENERATOR WITH EMPTY ROW STOPPING
function generateMessagesForSheetWithResume(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`❌ Sheet ${sheetName} not found`);
    return { generated: 0, skipped: 0 };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`📝 Sheet ${sheetName} has no data rows`);
    return { generated: 0, skipped: 0 };
  }
  
  // GET COLUMN R & S DATA FIRST to check what's already filled
  const messageRangeR = sheet.getRange(2, 15, lastRow - 1, 1); // Column R
  const messageRangeS = sheet.getRange(2, 16, lastRow - 1, 1); // Column S
  const existingMessagesR = messageRangeR.getValues();
  const existingMessagesS = messageRangeS.getValues();
  
  // Setup headers for columns R and S if needed
  if (sheet.getRange(1, 15).getValue() === '') {
    sheet.getRange(1, 15).setValue('English Message');
  }
  if (sheet.getRange(1, 16).getValue() === '') {
    sheet.getRange(1, 16).setValue('Arabic Message');
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 21); // Skip header, get A-U
  const data = dataRange.getValues();
  
  let generated = 0;
  let skipped = 0;
  let quotaHitThisRun = false;
  
  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const rowNumber = i + 2; // Actual row in sheet
    
    // 🛑 STOP if the entire row is empty (no more data below)
    if (isRowEmpty(rowData)) {
      Logger.log(`⏹️ Empty row detected at row ${rowNumber}. Stopping processing for sheet "${sheetName}".`);
      break;
    }
    
    const businessName = rowData[0]; // Column A - Name
    const address = rowData[1]; // Column B - Address
    const phone = rowData[2]; // Column C - Phone
    const website = rowData[3]; // Column D - Website
    const rating = rowData[5]; // Column F - Rating
    const totalReviews = rowData[6]; // Column G - Total Reviews
    const status = rowData[8]; // Column N - Status
    const analysis = rowData[14]; // Column Q - Website Analysis
    const demoWebsiteURL = rowData[17]; // Column U - Demo Website URL
    
    // CHECK IF ALREADY PROCESSED - Skip if both columns R and S have content
    const existingEnglish = existingMessagesR[i][0];
    const existingArabic = existingMessagesS[i][0];
    
    if (existingEnglish && existingEnglish !== "" && existingArabic && existingArabic !== "") {
      skipped++;
      continue; // Skip this row, already processed
    }
    
    // Skip rows without phone number
    if (!businessName || !phone) {
      skipped++;
      Logger.log(`⏭️ Skipping ${businessName || "row " + rowNumber} - no phone number`);
      continue;
    }
    
    const hasWebsite = website && website !== "" && website !== "No Website";
    const quotaExhaustedToday = isQuotaExhaustedToday();
    
    // If quota is exhausted today and this row needs a website analysis, skip it (don't generate message)
    if (hasWebsite && quotaExhaustedToday) {
      Logger.log(`⏸️ Quota exhausted today, skipping ${businessName} (needs PageSpeed). Will retry later.`);
      skipped++;
      continue;
    }
    
    Logger.log(`\n🔄 Processing ${sheetName} row ${rowNumber}: ${businessName}`);
    
    // GET PAGE SPEED SCORE WITH CACHE AND QUOTA HANDLING
    let pageSpeedData = null;
    
    if (hasWebsite && !urlFetchQuotaExhausted) {
      Logger.log(`🌐 Analyzing website: ${website}`);
      pageSpeedData = getConservativePageSpeedScore(website);
      
      if (pageSpeedData) {
        Logger.log(`✅ Got scores - Mobile: ${pageSpeedData.mobileScore}/100, Desktop: ${pageSpeedData.desktopScore}/100, Using: ${pageSpeedData.score}/100`);
      } else if (urlFetchQuotaExhausted) {
        Logger.log(`⚠️ PageSpeed API quota exhausted during this run. Marking today as exhausted.`);
        setQuotaExhaustedToday();
        quotaHitThisRun = true;
        // Skip this row because we couldn't get score and we want to retry later
        Logger.log(`⏸️ Skipping ${businessName} (quota exhausted). Will retry later.`);
        skipped++;
        continue;
      } else {
        Logger.log(`⚠️ Could not get PageSpeed score for ${website}`);
      }
    } else if (hasWebsite && urlFetchQuotaExhausted) {
      Logger.log(`⏩ Skipping PageSpeed for ${website} (quota exhausted this run)`);
    } else {
      Logger.log(`📭 No website available for ${businessName}`);
    }
    
    // Generate English message
    const englishMessage = createCustomMessageComplete({
      businessName: businessName,
      address: address,
      phone: phone,
      website: website,
      rating: rating,
      totalReviews: totalReviews,
      status: status,
      analysis: analysis,
      demoWebsiteURL: demoWebsiteURL,
      pageSpeedData: pageSpeedData
    });
    
    // Generate Arabic message
    const arabicMessage = createCustomMessageArabicComplete({
      businessName: businessName,
      address: address,
      phone: phone,
      website: website,
      rating: rating,
      totalReviews: totalReviews,
      status: status,
      analysis: analysis,
      demoWebsiteURL: demoWebsiteURL,
      pageSpeedData: pageSpeedData
    });
    
    // Write to columns R and S
    sheet.getRange(rowNumber, 15).setValue(englishMessage); // Column R
    sheet.getRange(rowNumber, 16).setValue(arabicMessage);  // Column S
    
    generated++;
    SpreadsheetApp.flush(); // Save progress
    
    Logger.log(`✅ Generated messages for ${businessName}`);
    
    // Very short pause every 10 rows to prevent script timeout
    if (generated > 0 && generated % 10 === 0) {
      Utilities.sleep(500);
    }
  }
  
  // Auto-resize columns for readability
  sheet.autoResizeColumns(15, 2);
  
  Logger.log(`✅ ${sheetName}: ${generated} generated, ${skipped} skipped`);
  return { generated: generated, skipped: skipped };
}

// Helper function to detect empty rows
function isRowEmpty(row) {
  return row.every(cell => cell === "" || cell === null || cell === undefined || (typeof cell === 'string' && cell.trim() === ""));
}

// FORMAT REVIEWS COUNT TO "2400+" FORMAT
function formatReviewsCount(reviews) {
  if (!reviews || isNaN(reviews)) return "";
  
  const reviewCount = parseInt(reviews);
  if (reviewCount < 10) return reviewCount.toString();
  
  // Round to nearest 10, 100, or 1000 based on size
  if (reviewCount < 100) {
    return Math.floor(reviewCount / 10) * 10 + "+";
  } else if (reviewCount < 1000) {
    return Math.floor(reviewCount / 100) * 100 + "+";
  } else {
    return Math.floor(reviewCount / 1000) * 1000 + "+";
  }
}

// SINGLE SCORE GETTER WITH QUOTA DETECTION
function getSinglePageSpeedScore(url, strategy = 'mobile') {
  if (urlFetchQuotaExhausted) return null;
  
  try {
    const apiUrl = CONFIG.pagespeedApiKey ? 
      `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encodeURIComponent(url)}&strategy=${strategy}&key=${CONFIG.pagespeedApiKey}` :
      `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encodeURIComponent(url)}&strategy=${strategy}`;
    
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true,
      timeout: 15000  // 15 seconds timeout
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const performanceScore = data.lighthouseResult?.categories?.performance?.score;
      if (performanceScore !== undefined && performanceScore !== null) {
        const roundedScore = Math.round(performanceScore * 100);
        Logger.log(`   ${strategy.toUpperCase()} Performance: ${roundedScore}/100`);
        return roundedScore;
      }
    } else if (responseCode === 429 || responseCode === 403) {
      Logger.log(`⚠️ PageSpeed API quota exceeded (HTTP ${responseCode}).`);
      urlFetchQuotaExhausted = true;
      return null;
    } else {
      Logger.log(`   ⚠️ ${strategy} returned HTTP ${responseCode}`);
      return null;
    }
  } catch (error) {
    const errorMsg = error.toString();
    if (errorMsg.includes('Service invoked too many times') || errorMsg.includes('quota')) {
      Logger.log(`⚠️ URLFetch daily limit reached.`);
      urlFetchQuotaExhausted = true;
    } else {
      Logger.log(`   ❌ ${strategy} test failed: ${errorMsg.substring(0, 100)}`);
    }
    return null;
  }
  return null;
}

// GET BOTH MOBILE AND DESKTOP, USE LOWER SCORE (CONSERVATIVE) WITH CACHING
function getConservativePageSpeedScore(url) {
  if (!url || url === "" || url === "No Website") return null;
  if (urlFetchQuotaExhausted) return null;
  
  let cleanUrl = url.trim();
  if (!cleanUrl.startsWith('http')) {
    cleanUrl = 'https://' + cleanUrl;
  }
  
  // Check cache first
  if (pageSpeedCache[cleanUrl]) {
    Logger.log(`🔄 Using cached score for ${cleanUrl}: ${pageSpeedCache[cleanUrl].score}/100`);
    return pageSpeedCache[cleanUrl];
  }
  
  Logger.log(`🔍 Testing: ${cleanUrl}`);
  
  // Get both scores (no sleep between)
  const mobileScore = getSinglePageSpeedScore(cleanUrl, 'mobile');
  const desktopScore = getSinglePageSpeedScore(cleanUrl, 'desktop');
  
  let result = null;
  
  if (mobileScore !== null && desktopScore !== null) {
    const conservativeScore = Math.min(mobileScore, desktopScore);
    Logger.log(`📊 Final Performance Scores - Mobile: ${mobileScore}/100, Desktop: ${desktopScore}/100, Using: ${conservativeScore}/100`);
    result = {
      score: conservativeScore,
      mobileScore: mobileScore,
      desktopScore: desktopScore,
      comment: getPerformanceComment(conservativeScore),
      url: cleanUrl,
      timestamp: new Date().toISOString(),
      category: 'performance',
      isConservative: true
    };
  } else if (mobileScore !== null) {
    Logger.log(`📊 Using mobile score only: ${mobileScore}/100`);
    result = {
      score: mobileScore,
      mobileScore: mobileScore,
      desktopScore: null,
      comment: getPerformanceComment(mobileScore),
      url: cleanUrl,
      timestamp: new Date().toISOString(),
      category: 'performance',
      isConservative: false
    };
  } else if (desktopScore !== null) {
    Logger.log(`📊 Using desktop score only: ${desktopScore}/100`);
    result = {
      score: desktopScore,
      mobileScore: null,
      desktopScore: desktopScore,
      comment: getPerformanceComment(desktopScore),
      url: cleanUrl,
      timestamp: new Date().toISOString(),
      category: 'performance',
      isConservative: false
    };
  }
  
  // Store in cache if we got any result
  if (result) {
    pageSpeedCache[cleanUrl] = result;
  }
  
  return result;
}

// ACCURATE PERFORMANCE COMMENTARY
function getPerformanceComment(score) {
  if (score < 50) {
    return {
      english: "which indicates poor performance and urgent need for speed optimization",
      arabic: "مما يشير إلى أداء ضعيف وحاجة ملحة لتحسين السرعة"
    };
  } else if (score < 70) {
    return {
      english: "which indicates below average performance with significant room for improvement, especially on mobile",
      arabic: "مما يشير إلى أداء أقل من المتوسط مع مجال كبير للتحسين، خصوصاً على الجوال"
    };
  } else if (score < 85) {
    return {
      english: "which indicates average performance with opportunities for optimization to improve user experience",
      arabic: "مما يشير إلى أداء متوسط مع فرص للتحسين لتعزيز تجربة المستخدم"
    };
  } else if (score < 95) {
    return {
      english: "which is good performance but can be further optimized for the best user experience",
      arabic: "وهو أداء جيد ولكن يمكن تحسينه أكثر لأفضل تجربة مستخدم"
    };
  } else {
    return {
      english: "which is excellent performance - maintaining this ensures optimal user experience",
      arabic: "وهو أداء ممتاز - الحفاظ على هذا المستوى يضمن تجربة مستخدم مثالية"
    };
  }
}

// ENHANCED ENGLISH MESSAGE WITH EXACT FORMAT FROM EXAMPLE
function createCustomMessageComplete(businessData) {
  const { businessName, address, phone, website, rating, totalReviews, status, analysis, demoWebsiteURL, pageSpeedData } = businessData;
  
  const cleanName = cleanBusinessName(businessName);
  const area = extractNeighborhoodFromAddressEnglish(address);
  const hasWebsite = website && website !== "" && website !== "No Website";
  const hasAnalysis = analysis && analysis.includes("OPPORTUNITIES");
  
  // Format reviews count
  const formattedReviews = formatReviewsCount(totalReviews);
  
  // Build professional English message
  let message = `Dear ${cleanName} Team,\n\n`;
  message += `I'm ${CONFIG.englishName}, a data analyst at ${CONFIG.companyNameEnglish}. `;
  
  // Add personalized details with rating and area
  if (rating && totalReviews) {
    message += `I reviewed your Google Business Profile and noticed your excellent ${rating}-star rating with ${formattedReviews} reviews 📊`;
    if (area) {
      message += ` in ${area} district`;
    }
    message += ` — impressive results that show customer trust in your services. `;
  } else if (area) {
    message += `I see your business is located in ${area} district. `;
  }
  
  // Add PageSpeed analysis if available (matching example format)
  if (hasWebsite && pageSpeedData !== null) {
    if (pageSpeedData.mobileScore !== null && pageSpeedData.desktopScore !== null) {
      message += `\n\nWhen analyzing your website's mobile performance, the score was ${pageSpeedData.mobileScore}/100 according to Google Lighthouse (Desktop: ${pageSpeedData.desktopScore}/100),\n${pageSpeedData.comment.english}. `;
    } else if (pageSpeedData.mobileScore !== null) {
      message += `\n\nWhen analyzing your website's mobile performance, the score was ${pageSpeedData.mobileScore}/100 according to Google Lighthouse,\n${pageSpeedData.comment.english}. `;
    } else if (pageSpeedData.desktopScore !== null) {
      message += `\n\nWhen analyzing your website's desktop performance, the score was ${pageSpeedData.desktopScore}/100 according to Google Lighthouse,\n${pageSpeedData.comment.english}. `;
    }
  }
  
  // Handle different scenarios
  if (!hasWebsite) {
    message += `\n\nHowever, you're missing significant growth opportunities by not having a professional website.\n\n`;
  } else if (hasAnalysis) {
    const topOpportunities = extractTopOpportunitiesFromAnalysis(analysis, 2);
    if (topOpportunities.length > 0) {
      message += `\n\nBased on our analysis, we found key areas for improvement.\n\n`;
    } else {
      message += `\n\nBased on our expertise, we can help you enhance your digital presence.\n\n`;
    }
  } else {
    message += `\n\nBased on our expertise, we can help you optimize your digital performance.\n\n`;
  }
  
  // Value proposition
  message += `Through our experience, we can assist you with:\n`;
  
  if (!hasWebsite) {
    message += `• Building a professional, mobile-responsive website\n`;
    message += `• Implementing analytics and booking tracking systems\n`;
    message += `• Improving search engine visibility\n`;
    message += `• Developing a complete online reservation system\n`;
    message += `• Setting up dynamic remarketing to re-engage visitors\n`;
  } else {
    // Always include speed optimization if score is low
    if (pageSpeedData !== null && pageSpeedData.score < 75) {
      message += `• Improving website speed and visitor experience\n`;
    }
    
    // Include analytics if missing - with dynamic remarketing
    if (hasAnalysis && analysis.includes('analytics') && !analysis.includes('GA4 installed')) {
      message += `• Setting up analytics and conversion tracking\n`;
      message += `• Implementing dynamic remarketing to target interested visitors\n`;
    }
    
    // Include SEO if missing
    if (hasAnalysis && analysis.includes('SEO')) {
      message += `• Boosting your visibility in search results\n`;
    }
    
    // Business-specific features
    const businessType = detectBusinessType(cleanName);
    const specificBenefits = getBusinessSpecificSolutions(businessType, hasWebsite);
    specificBenefits.forEach(benefit => {
      message += `• ${benefit}\n`;
    });
    
    // Fallback if no specific benefits
    if (specificBenefits.length === 0) {
      message += `• Optimizing website performance and user experience\n`;
      message += `• Implementing proper analytics and tracking\n`;
      message += `• Enhancing search engine visibility\n`;
      message += `• Setting up audience targeting and remarketing campaigns\n`;
    }
  }
  
  message += `\n`;
  
  // Add demo website section
  if (demoWebsiteURL && demoWebsiteURL !== "") {
    message += `We've prepared a demo website showcasing our capabilities — this is an initial version that we can fully customize to meet your specific needs.\n`;
    message += `${demoWebsiteURL}\n\n`;
  }
  
  // Professional closing
  message += `Would a brief call in the coming days work to discuss how we can enhance your digital presence? ☎️\n\n`;
  
  message += `Best regards,\n`;
  message += `${CONFIG.englishName}\n`;
  message += `${CONFIG.companyNameEnglish}\n`;
  message += `📱 ${CONFIG.phoneNumber}`;
  
  return message;
}

// ENHANCED ARABIC MESSAGE WITH EXACT FORMAT FROM EXAMPLE
function createCustomMessageArabicComplete(businessData) {
  const { businessName, address, phone, website, rating, totalReviews, status, analysis, demoWebsiteURL, pageSpeedData } = businessData;
  
  const cleanName = cleanBusinessNameArabic(businessName);
  const area = extractNeighborhoodFromAddressArabic(address);
  const hasWebsite = website && website !== "" && website !== "No Website";
  const hasAnalysis = analysis && analysis.includes("OPPORTUNITIES");
  
  // Format rating and reviews
  const formattedRating = rating ? rating.toString().replace('.', ',') : "ممتاز";
  const formattedReviews = formatReviewsCount(totalReviews);
  
  // Build natural Arabic message WITH COMPLETE DIRECTION CONTROL
  let message = `\u202B`; // Right-to-Left Embed (RLE) character at start
  
  message += `السلام عليكم،\n\n`;
  message += `أنا ${CONFIG.arabicName}، محلل بيانات في شركة ${CONFIG.companyName}. `;
  
  // Add personalized details with rating and area
  if (rating && totalReviews) {
    message += `اطلعت على صفحة ${cleanName} على جوجل ولاحظت تقييمكم المميز (${formattedRating} نجوم ⭐ من أكثر من ${formattedReviews} تقييم)`;
    if (area) {
      message += ` في حي ${area}`;
    }
    message += ` — ما شاء الله، دليل على ثقة الزوار بخدماتكم. `;
  } else if (area) {
    message += `أرى أن متجركم يقع في حي ${area}. `;
  }
  
  // Add PageSpeed analysis if available (matching example format)
  if (hasWebsite && pageSpeedData !== null) {
    if (pageSpeedData.mobileScore !== null && pageSpeedData.desktopScore !== null) {
      message += `\n\nعند تحليل موقعكم الإلكتروني، كانت نتيجة الجوال ${pageSpeedData.mobileScore}/100 وسطح المكتب ${pageSpeedData.desktopScore}/100 وفق جوجل لايت هاوس,\n${pageSpeedData.comment.arabic}. `;
    } else if (pageSpeedData.mobileScore !== null) {
      message += `\n\nعند تحليل موقعكم الإلكتروني، كانت نتيجة الجوال ${pageSpeedData.mobileScore}/100 وفق جوجل لايت هاوس,\n${pageSpeedData.comment.arabic}. `;
    } else if (pageSpeedData.desktopScore !== null) {
      message += `\n\nعند تحليل موقعكم الإلكتروني، كانت نتيجة سطح المكتب ${pageSpeedData.desktopScore}/100 وفق جوجل لايت هاوس,\n${pageSpeedData.comment.arabic}. `;
    }
  }

  // Handle different scenarios
  if (!hasWebsite) {
    message += `\n\nمن خلال خبرتنا، يمكننا مساعدتكم في تعزيز حضوركم الرقمي.\n\n`;
  } else if (hasAnalysis) {
    const topOpportunities = extractTopOpportunitiesFromAnalysisArabic(analysis, 2);
    if (topOpportunities.length > 0) {
      message += `\n\nمن خلال تحليلنا، وجدنا نقاط رئيسية للتحسين.\n\n`;
    } else {
      message += `\n\nمن خلال خبرتنا، يمكننا مساعدتكم في تعزيز حضوركم الرقمي.\n\n`;
    }
  } else {
    message += `\n\nمن خلال خبرتنا، يمكننا مساعدتكم في تحسين أدائكم الرقمي.\n\n`;
  }
  
  // Value proposition
  message += ` يمكننا مساعدتكم في:\n`;
  
  if (!hasWebsite) {
    message += `\u2022 بناء موقع إلكتروني احترافي متجاوب مع الجوال\n`;
    message += `\u2022 إعداد أنظمة التحليل وتتبع الحجوزات\n`;
    message += `\u2022 رفع الظهور في نتائج البحث\n`;
    message += `\u2022 تطوير نظام حجز إلكتروني متكامل\n`;
    message += `\u2022 إعداد إعلانات ديناميكية لإعادة استهداف الزوار\n`;
  } else {
    // Always include speed optimization if score is low
    if (pageSpeedData !== null && pageSpeedData.score < 75) {
      message += `\u2022 تحسين سرعة الموقع وتجربة الزوار\n`;
    }
    
    // Include analytics if missing - with dynamic remarketing
    if (hasAnalysis && analysis.includes('analytics') && !analysis.includes('GA4 installed')) {
      message += `\u2022 إعداد أنظمة التحليل وتتبع التحويلات\n`;
      message += `\u2022 تطبيق الإعلانات الديناميكية لإعادة استهداف الزوار المهتمين\n`;
    }
    
    // Include SEO if missing
    if (hasAnalysis && analysis.includes('SEO')) {
      message += `\u2022 رفع الظهور في نتائج البحث\n`;
    }
    
    // Business-specific features
    const businessType = detectBusinessTypeArabic(cleanName);
    const specificBenefits = getBusinessSpecificSolutionsArabic(businessType, hasWebsite);
    specificBenefits.forEach(benefit => {
      message += `\u2022 ${benefit}\n`;
    });
    
    // Fallback if no specific benefits
    if (specificBenefits.length === 0) {
      message += `\u2022 تحسين أداء الموقع وتجربة المستخدم\n`;
      message += `\u2022 تطبيق أنظمة التحليل والتتبع\n`;
      message += `\u2022 تعزيز الظهور في محركات البحث\n`;
      message += `\u2022 إعداد حملات استهداف الجمهور والإعلانات الاسترجاعية\n`;
    }
  }
  
  message += `\n`;
  
  // Add demo website section
  if (demoWebsiteURL && demoWebsiteURL !== "") {
    message += `قمنا بإعداد موقع تجريبي لعرض إمكانياتنا — هذه نسخة أولية قابلة للتخصيص الكامل بما يناسب احتياجاتكم.\n`;
    message += `${demoWebsiteURL}\n\n`;
  }
  
  // Natural Arabic closing
  message += `هل يناسبكم مكالمة قصيرة خلال الأيام القادمة لمناقشة كيف يمكننا تعزيز حضوركم الرقمي؟ ☎️\n\n`;
  
  message += `مع خالص التحية،\n`;
  message += `${CONFIG.arabicName}\n`;
  message += `${CONFIG.companyName}\n`;
  message += `\u202A📱 ${CONFIG.phoneNumber}\u202C`; // LRE for phone number, then PDF
  
  // Add PDF (Pop Directional Formatting) to end the RTL context
  message += `\u202C`;
  
  return message;
}

// OPTIONAL: FUNCTION TO COPY ARABIC TEXT WITH PROPER FORMATTING
function copyArabicToClipboard(sheetName, rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;
  
  const arabicMessage = sheet.getRange(rowNumber, 16).getValue(); // Column P (Arabic)
  
  // Create a temporary HTML service to copy formatted text
  const htmlOutput = HtmlService.createHtmlOutput(`
    <textarea id="arabicText" style="width: 100%; height: 300px; direction: rtl; text-align: right;" dir="rtl">${arabicMessage}</textarea>
    <br>
    <button onclick="copyText()">نسخ النص</button>
    <script>
      function copyText() {
        const textarea = document.getElementById('arabicText');
        textarea.select();
        document.execCommand('copy');
        alert('تم نسخ النص بنجاح!');
      }
    </script>
  `)
  .setWidth(600)
  .setHeight(400)
  .setTitle('نسخ النص العربي');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'نسخ النص العربي');
  
  return arabicMessage;
}

// BUSINESS-SPECIFIC SOLUTIONS FOR ENGLISH
function getBusinessSpecificSolutions(businessType, hasWebsite) {
  const solutions = {
    'Hotel': [
      'Developing a complete online booking system',
      'Interactive room displays and special offers',
      'Optimizing for local search and tourism keywords'
    ],
    'Restaurant': [
      'Online menu with high-quality images',
      'Table reservation and online ordering system',
      'Showcasing daily specials and promotions'
    ],
    'Car Dealer': [
      'Vehicle showcase with detailed specifications',
      'Online test drive booking system',
      'Maintenance appointment scheduling'
    ],
    'Real Estate': [
      'Property listings with advanced filters',
      'Virtual tours and image galleries',
      'Online appointment scheduling for viewings'
    ],
    'Pharmacy': [
      'Online appointment booking system',
      'Product catalog with search functionality',
      'Prescription management features'
    ],
    'Electronics': [
      'Professional product display and catalog',
      'Online reservation for products/services',
      'Technical specifications and comparisons'
    ],
    'Supermarket': [
      'Product catalog with categories',
      'Weekly offers and delivery scheduling',
      'Online ordering system'
    ],
    'Beauty Salon': [
      'Service menu with online booking',
      'Stylist portfolio and reviews',
      'Special packages and promotions'
    ],
    'Gym': [
      'Membership plans and online registration',
      'Class schedule and booking system',
      'Trainer profiles and specialties'
    ],
    'Cafe': [
      'Interactive menu with high-quality images',
      'Online ordering and reservation system',
      'Special events and loyalty programs'
    ]
  };
  
  return solutions[businessType] || [
    'Professional online presence optimization',
    'Customer engagement and lead generation',
    'Digital marketing strategy development'
  ];
}

// BUSINESS-SPECIFIC SOLUTIONS FOR ARABIC
function getBusinessSpecificSolutionsArabic(businessType, hasWebsite) {
  const solutions = {
    'فندق': [
      'تطوير نظام حجز إلكتروني متكامل',
      'عرض الغرف والعروض الخاصة بطريقة تفاعلية',
      'تحسين الظهور في البحث المحلي والسياحي'
    ],
    'مطعم': [
      'قائمة طعام إلكترونية مع صور عالية الجودة',
      'نظام حجز الطاولات والطلب أونلاين',
      'عرض العروض اليومية والترويجات'
    ],
    'سيارات': [
      'عرض السيارات بمواصفات مفصلة',
      'نظام حجز تجربة قيادة إلكتروني',
      'جدولة مواعيد الصيانة'
    ],
    'عقارات': [
      'قوائم العقارات مع فلاتر متقدمة',
      'جولات افتراضية ومعارض الصور',
      'جدولة مواعيد المشاهدة أونلاين'
    ],
    'صيدلية': [
      'نظام حجز مواعيد إلكتروني',
      'كتالوج المنتجات مع خاصية البحث',
      'إدارة الوصفات الطبية'
    ],
    'إلكترونيات': [
      'عرض المنتجات بشكل احترافي',
      'نظام حجز إلكتروني للمنتجات والخدمات',
      'المواصفات التقنية والمقارنات'
    ],
    'سوبرماركت': [
      'كتالوج المنتجات مع التصنيفات',
      'العروض الأسبوعية وجدولة التوصيل',
      'نظام الطلب الإلكتروني'
    ],
    'صالون تجميل': [
      'قائمة الخدمات مع الحجز أونلاين',
      'معرض أعمال المصممين وتقييماتهم',
      'الباقات الخاصة والترويجات'
    ],
    'نادي رياضي': [
      'باقات العضوية والتسجيل الإلكتروني',
      'جدول الحصص ونظام الحجز',
      'ملفات المدربين وتخصصاتهم'
    ],
    'مقهى': [
      'قائمة تفاعلية مع صور عالية الجودة',
      'نظام الطلب والحجز الإلكتروني',
      'الفعاليات الخاصة وبرامج الولاء'
    ]
  };
  
  return solutions[businessType] || [
    'تحسين الحضور الإلكتروني الاحترافي',
    'تفاعل العملاء واستقطابهم',
    'تطوير استراتيجية التسويق الرقمي'
  ];
}

// TESTING FUNCTIONS
function testPageSpeedForUrl() {
  const testUrl = "https://www.gloriainn.sa"; // Change this to test any URL
  Logger.log(`🧪 Testing PageSpeed for: ${testUrl}`);
  
  const result = getConservativePageSpeedScore(testUrl);
  if (result) {
    Logger.log(`✅ Final Score: ${result.score}/100`);
    Logger.log(`✅ Mobile: ${result.mobileScore}/100`);
    Logger.log(`✅ Desktop: ${result.desktopScore}/100`);
  } else {
    Logger.log(`❌ Failed to get PageSpeed data`);
  }
  
  return result;
}

// SUPPORTING FUNCTIONS

function cleanBusinessName(businessName) {
  if (!businessName) return "";
  let clean = businessName
    .replace(/\s*(LLC|L\.L\.C|Inc|Incorporated|Corp|Corporation|Ltd|Limited|\.com|\.net)\s*$/gi, '')
    .replace(/\s*[,.-]\s*$/, '')
    .trim();
  return clean || "Business";
}

function cleanBusinessNameArabic(businessName) {
  if (!businessName) return "";
  let clean = businessName
    .replace(/\s*(LLC|L\.L\.C|Inc|Incorporated|Corp|Corporation|Ltd|Limited|\.com|\.net)\s*$/gi, '')
    .replace(/\s*[,.-]\s*$/, '')
    .trim();
  
  if (!containsArabic(clean)) {
    const translations = {
      'electronics': 'الإلكترونيات', 'pharmacy': 'الصيدلية', 'restaurant': 'المطعم',
      'hotel': 'الفندق', 'supermarket': 'السوبرماركت', 'gym': 'النادي الرياضي',
      'cafe': 'المقهى', 'car': 'السيارات', 'real estate': 'العقارات',
      'beauty': 'صالون التجميل'
    };
    for (const [eng, arb] of Object.entries(translations)) {
      if (clean.toLowerCase().includes(eng)) return arb;
    }
  }
  return clean || "العميل";
}

function extractNeighborhoodFromAddressEnglish(address) {
  if (!address) return null;
  
  const neighborhoodTranslations = {
    'ظهرة لبن': 'Dahrat Laban', 'وادي لبن': 'Wadi Laban', 'الملز': 'Al Malaz', 
    'العليا': 'Al Olaya', 'الرياض': 'Riyadh', 'النخيل': 'Al Nakhil',
    'العارض': 'Al Arad', 'المرسلات': 'Al Mursalat', 'العريجاء': 'Al Uraija', 
    'السويدي': 'Al Suwaidi', 'طويق': 'Tuwaiq', 'الشفا': 'Al Shifa',
    'الزهراء': 'Al Zahra', 'الربوة': 'Al Rabwa', 'الخالدية': 'Al Khalidiyyah', 
    'العزيزية': 'Al Aziziyah', 'المنصورة': 'Al Mansourah',
    'الفيصلية': 'Al Faisaliyyah', 'الرحمانية': 'Al Rahmaniyyah', 'حطين': 'Hittin', 
    'النهضة': 'Al Nahdah', 'الروضة': 'Al Rawdah',
    'الوزارات': 'Al Wizarat', 'الشميسي': 'Al Shemaysi', 'البطحاء': 'Al Bathaa', 
    'ثليم': 'Thulaim', 'القدس': 'Al Quds', 'السلام': 'Al Salam',
    'الاندلس': 'Al Andalus', 'الفيحاء': 'Al Fayha', 'الروابي': 'Al Rawabi', 
    'الجزيرة': 'Al Jazirah', 'المروج': 'Al Muruj', 'المصانع': 'Al Masani',
    'الوشام': 'Al Wisham', 'السلطانة': 'Al Sultanah', 'البديعة': 'Al Badi ah'
  };

  const arabicNeighborhood = extractNeighborhoodFromAddressArabic(address);
  if (arabicNeighborhood) {
    return neighborhoodTranslations[arabicNeighborhood] || arabicNeighborhood;
  }
  
  const addressLower = address.toLowerCase();
  for (const [arabic, english] of Object.entries(neighborhoodTranslations)) {
    if (addressLower.includes(english.toLowerCase())) {
      return english;
    }
  }
  
  return null;
}

function extractNeighborhoodFromAddressArabic(address) {
  if (!address) return null;
  const neighborhoodTranslations = {
    'Dahrat Laban': 'ظهرة لبن', 'Wadi Laban': 'وادي لبن', 'Al Malaz': 'الملز', 'Al Olaya': 'العليا', 'Riyadh': 'الرياض',
    'Al Nakhil': 'النخيل', 'Al Arad': 'العارض', 'Al Mursalat': 'المرسلات', 'Al Uraija': 'العريجاء', 'Al Suwaidi': 'السويدي',
    'Tuwaiq': 'طويق', 'Al Shifa': 'الشفا', 'Al Zahra': 'الزهراء', 'Al Rabwa': 'الربوة', 'Al Khalidiyyah': 'الخالدية',
    'Al Aziziyah': 'العزيزية', 'Al Mansourah': 'المنصورة', 'Al Faisaliyyah': 'الفيصلية', 'Al Rahmaniyyah': 'الرحمانية',
    'Hittin': 'حطين', 'Al Nahdah': 'النهضة', 'Al Rawdah': 'الروضة', 'Al Wizarat': 'الوزارات', 'Al Shemaysi': 'الشميسي',
    'Al Bathaa': 'البطحاء', 'Thulaim': 'ثليم', 'Al Quds': 'القدس', 'Al Salam': 'السلام', 'Al Andalus': 'الاندلس',
    'Al Fayha': 'الفيحاء', 'Al Rawabi': 'الروابي', 'Al Jazirah': 'الجزيرة', 'Al Muruj': 'المروج', 'Al Masani': 'المصانع',
    'Al Wisham': 'الوشام', 'Al Sultanah': 'السلطانة', 'Al Badi ah': 'البديعة'
  };
  const addressLower = address.toLowerCase();
  const arabicNeighborhoods = [
    'ظهرة لبن', 'وادي لبن', 'الملز', 'العليا', 'الرياض', 'النخيل', 'العارض', 'المرسلات', 'العريجاء', 'السويدي', 'طويق', 'الشفا', 'الزهراء',
    'الربوة', 'الخالدية', 'العزيزية', 'المنصورة', 'الفيصلية', 'الرحمانية', 'حطين', 'النهضة', 'الروضة', 'الوزارات', 'الشميسي', 'البطحاء',
    'ثليم', 'القدس', 'السلام', 'الاندلس', 'الفيحاء', 'الروابي', 'الجزيرة', 'المروج', 'المصانع', 'الوشام', 'السلطانة', 'البديعة', 'السفارات',
    'عرقة', 'المهندسين', 'المربع', 'الضباط', 'المرور', 'الطيران', 'الدخل', 'الفلاح', 'النسيم', 'الخيران'
  ];
  for (const neighborhood of arabicNeighborhoods) {
    if (address.includes(neighborhood)) {
      return neighborhood;
    }
  }
  for (const [englishName, arabicName] of Object.entries(neighborhoodTranslations)) {
    if (addressLower.includes(englishName.toLowerCase())) {
      return arabicName;
    }
  }
  const patterns = [ /حي\s*([^,،]+)/, /في\s*([^,،]+)/, /ب\s*([^,،]+)/, /,\s*([^,،]+),/ ];
  for (const pattern of patterns) {
    const match = address.match(pattern);
    if (match && match[1]) {
      const potentialArea = match[1].trim();
      if (potentialArea.length > 3 && potentialArea.length < 20) {
        return potentialArea;
      }
    }
  }
  return null;
}

function detectBusinessType(businessName) {
  const nameLower = businessName.toLowerCase();
  if (nameLower.includes('electron') || nameIncludesArabicElectronics(businessName)) return 'Electronics';
  if (nameLower.includes('pharm')) return 'Pharmacy';
  if (nameLower.includes('car')) return 'Car Dealer';
  if (nameLower.includes('real')) return 'Real Estate';
  if (nameLower.includes('hotel')) return 'Hotel';
  if (nameLower.includes('super')) return 'Supermarket';
  if (nameLower.includes('beauty')) return 'Beauty Salon';
  if (nameLower.includes('cafe')) return 'Cafe';
  if (nameLower.includes('gym')) return 'Gym';
  if (nameLower.includes('restaurant')) return 'Restaurant';
  return 'Store';
}

function detectBusinessTypeArabic(businessName) {
  const nameLower = businessName.toLowerCase();
  if (nameLower.includes('electron') || nameIncludesArabicElectronics(businessName)) return 'إلكترونيات';
  if (nameLower.includes('pharm') || nameLower.includes('صيد')) return 'صيدلية';
  if (nameLower.includes('car') || nameLower.includes('سيار')) return 'سيارات';
  if (nameLower.includes('real') || nameLower.includes('عقار')) return 'عقارات';
  if (nameLower.includes('hotel') || nameLower.includes('فندق')) return 'فندق';
  if (nameLower.includes('super') || nameLower.includes('سوبر')) return 'سوبرماركت';
  if (nameLower.includes('beauty') || nameLower.includes('جمال')) return 'صالون تجميل';
  if (nameLower.includes('cafe') || nameLower.includes('كاف')) return 'مقهى';
  if (nameLower.includes('gym') || nameLower.includes('جيم')) return 'نادي رياضي';
  if (nameLower.includes('restaurant') || nameLower.includes('مطعم')) return 'مطعم';
  return 'متجر';
}

function nameIncludesArabicElectronics(businessName) {
  const arabicElectronicsWords = ['إلكترونيات', 'كهر', 'كهرب', 'أجهزة', 'تلفون', 'جوال', 'موبايل', 'لاب', 'كمبيوتر'];
  return arabicElectronicsWords.some(word => businessName.includes(word));
}

function containsArabic(text) {
  if (!text) return false;
  return /[\u0600-\u06FF]/.test(text);
}

function extractTopOpportunitiesFromAnalysis(analysis, count = 2) {
  const opportunities = [];
  
  if (!analysis || !analysis.includes("OPPORTUNITIES")) return opportunities;
  
  const lines = analysis.split('\n');
  let currentCategory = '';
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine || trimmedLine.includes('📊') || trimmedLine.includes('🛠') || trimmedLine.includes('✅')) continue;
    
    if (trimmedLine.includes('Analytics & Tracking')) currentCategory = 'Analytics';
    else if (trimmedLine.includes('SEO & Meta')) currentCategory = 'SEO';
    else if (trimmedLine.includes('Mobile & UX')) currentCategory = 'Mobile';
    else if (trimmedLine.includes('Business Tools')) currentCategory = 'Business';
    else if (trimmedLine.startsWith('•') && currentCategory) {
      const description = trimmedLine.substring(1).trim();
      opportunities.push({
        description: description,
        category: currentCategory,
        priority: getOpportunityPriority(description)
      });
    }
  }
  
  return opportunities.sort((a, b) => b.priority - a.priority).slice(0, count);
}

function extractTopOpportunitiesFromAnalysisArabic(analysis, count = 2) {
  const opportunities = [];
  
  if (!analysis || !analysis.includes("OPPORTUNITIES")) return opportunities;
  
  const lines = analysis.split('\n');
  let currentCategory = '';
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine || trimmedLine.includes('📊') || trimmedLine.includes('🛠') || trimmedLine.includes('✅')) continue;
    
    if (trimmedLine.includes('Analytics & Tracking')) currentCategory = 'Analytics';
    else if (trimmedLine.includes('SEO & Meta')) currentCategory = 'SEO';
    else if (trimmedLine.includes('Mobile & UX')) currentCategory = 'Mobile';
    else if (trimmedLine.includes('Business Tools')) currentCategory = 'Business';
    else if (trimmedLine.startsWith('•') && currentCategory) {
      const description = translateOpportunityToArabic(trimmedLine.substring(1).trim());
      opportunities.push({
        description: description,
        category: currentCategory,
        priority: getOpportunityPriority(trimmedLine.substring(1).trim())
      });
    }
  }
  
  return opportunities.sort((a, b) => b.priority - a.priority).slice(0, count);
}

function translateOpportunityToArabic(opportunity) {
  const translations = {
    'GA4 not installed': `نظام تحليل الزوار (Google Analytics 4) غير مثبت - يفقدكم بيانات مهمة عن زوار الموقع`,
    'Google Analytics not installed': `نظام تحليل الزوار غير مثبت - تفقدون بيانات مهمة عن زوار موقعكم`,
    'Using outdated Universal Analytics': `تستخدمون نظام تحليل قديم - يجب الترقية إلى Google Analytics 4`,
    'Google Tag Manager missing': `أداة إدارة الأكواد (Google Tag Manager) غير مثبتة - صعوبة في إدارة أدوات التتبع`,
    'Facebook Pixel not found': `أداة تتبع فيسبوك (Facebook Pixel) غير مثبتة - تفقدون تحليلات الحملات الإعلانية`,
    'Google Ads tracking not detected': `تتبع إعلانات جوجل غير مفعل - لا يمكن قياس نجاح الحملات الإعلانية`,
    'No heatmap analytics': `أدوات تحليل سلوك الزوار غير مثبتة - لا يمكن فهم كيفية تفاعل الزوار مع الموقع`,
    'Meta description missing': `الوصف التعريفي للموقع غير موجود - يؤثر على ظهوركم في نتائج البحث`,
    'Open Graph tags missing': `وسوم Open Graph مفقودة - تقليل جودة المشاركة على وسائل التواصل`,
    'Twitter Card tags missing': `وسوم Twitter Card غير موجودة - تقليل جودة المشاركة على تويتر`,
    'Structured data missing': `البيانات المنظمة غير مضبوطة - تفويت فرص الظهور في النتائج المميزة`,
    'No XML sitemap': `خريطة الموقع غير موجودة - صعوبة في فهرسة المحتوى بمحركات البحث`,
    'No SSL certificate': `شهادة SSL غير مفعلة - تأثير سلبي على الأمان وتصنيفات البحث`,
    'Mobile optimization issues': `مشاكل في تحسين الجوال - تجربة مستخدم غير مريحة على الهواتف`,
    'No viewport meta tag': `إعدادات الجوال غير مضبوطة - مشاكل في عرض الموقع على الهواتف`,
    'Touch targets too small': `أزرار الموقع صغيرة جداً - صعوبة في الاستخدام على الشاشات الصغيرة`,
    'Slow page speed': `سرعة الموقع بطيئة - تأثير سلبي على تجربة المستخدم والترتيب في البحث`,
    'No CDN implementation': `شبكة توصيل المحتوى غير مفعلة - بطء في التحميل للزوار من مناطق مختلفة`,
    'Contact form not detected': `نموذج الاتصال غير موجود - فقدان فرص استقبال استفسارات العملاء`,
    'Live chat not installed': `خدمة الدردشة المباشرة غير متوفرة - تفويت فرص البيع الفوري`,
    'No email subscription form': `نموذج الاشتراك في النشرة البريدية غير موجود - فقدان فرص التسويق المتكرر`,
    'No booking/appointment system': `نظام الحجوزات غير موجود - تعقيد في عملية حجز العملاء`,
    'No customer reviews system': `نظام التقييمات غير مفعل - فقدان عنصر الثقة والاجتماعي`,
    'No clear call-to-action buttons': `أزرار الحث على الإجراء غير واضحة - تقليل معدلات التحويل`,
    'No exit-intent popup': `نافذة الخروج غير مفعلة - فقدان فرص استعادة الزوار`,
    'No A/B testing setup': `اختبارات A/B غير مفعلة - صعوبة في تحسين أداء الموقع`,
    'No security headers': `إعدادات الأمان غير مكتملة - تعريض الموقع لمخاطر أمنية`,
    'Vulnerable libraries detected': `مكتبات قديمة وغير آمنة - خطر على أمان الموقع`,
    'Poor SEO optimization': `تحسين محركات البحث ضعيف - تقليل الظهور في النتائج`,
    'Missing lead capture': `نظام استقطاب العملاء غير فعال - فقدان فرص البيع`,
    'Missing conversion tracking': `تتبع التحويلات غير مفعل - عدم القدرة على قياس عائد الاستثمار`
  };
  
  let translated = opportunity;
  for (const [eng, arb] of Object.entries(translations)) {
    if (opportunity.includes(eng)) {
      translated = translated.replace(eng, arb);
      break;
    }
  }
  
  if (translated === opportunity) {
    if (opportunity.includes('analytics') || opportunity.includes('Analytics')) {
      translated = `مشاكل في أنظمة التحليل والتتبع - ${opportunity}`;
    } else if (opportunity.includes('SEO') || opportunity.includes('search')) {
      translated = `مشاكل في تحسين محركات البحث - ${opportunity}`;
    } else if (opportunity.includes('mobile') || opportunity.includes('Mobile')) {
      translated = `مشاكل في تجربة الجوال - ${opportunity}`;
    } else if (opportunity.includes('contact') || opportunity.includes('form')) {
      translated = `مشاكل في استقبال استفسارات العملاء - ${opportunity}`;
    }
  }
  
  return translated;
}

function getOpportunityPriority(opportunity) {
  if (opportunity.includes('GA4') || opportunity.includes('analytics')) return 10;
  if (opportunity.includes('Mobile') || opportunity.includes('optimization')) return 9;
  if (opportunity.includes('SEO') || opportunity.includes('Meta description')) return 8;
  if (opportunity.includes('Google Tag Manager')) return 7;
  if (opportunity.includes('Facebook Pixel') || opportunity.includes('Google Ads')) return 6;
  if (opportunity.includes('Contact form') || opportunity.includes('Live chat')) return 5;
  if (opportunity.includes('Structured data') || opportunity.includes('Open Graph')) return 4;
  return 3;
}

function parseAnalysisData(analysis) {
  if (!analysis) return { 
    opportunityCount: 0, 
    hasMissingAnalytics: false,
    hasMobileIssues: false,
    hasSEOMissing: false,
    hasConversionIssues: false
  };
  
  const opportunityMatch = analysis.match(/OPPORTUNITIES: (\d+)/);
  const opportunityCount = opportunityMatch ? parseInt(opportunityMatch[1]) : 0;
  
  const hasMissingAnalytics = /analytics|tracking|pixel|gtag|ga\(/i.test(analysis);
  const hasMobileIssues = /mobile|viewport|responsive/i.test(analysis);
  const hasSEOMissing = /SEO|meta|description|structured/i.test(analysis);
  const hasConversionIssues = /contact|form|chat|conversion/i.test(analysis);
  
  return {
    opportunityCount: opportunityCount,
    hasMissingAnalytics: hasMissingAnalytics,
    hasMobileIssues: hasMobileIssues,
    hasSEOMissing: hasSEOMissing,
    hasConversionIssues: hasConversionIssues
  };
}

function addBusinessSpecificBenefits(businessType, hasWebsite) {
  let benefits = '';
  
  const benefitMap = {
    'Electronics': [
      'Professional product display and catalog',
      'Online reservation system for products/services',
      'Technical specifications and comparisons'
    ],
    'Pharmacy': [
      'Online appointment booking system',
      'Product catalog with search functionality',
      'Prescription management features'
    ],
    'Car Dealer': [
      'Vehicle showcase with detailed specifications',
      'Maintenance appointment scheduling',
      'Test drive booking system'
    ],
    'Restaurant': [
      'Online menu with images',
      'Table reservation system',
      'Online ordering for takeaway/delivery'
    ],
    'Hotel': [
      'Online booking system',
      'Room gallery and amenities display',
      'Special offers and packages'
    ],
    'Supermarket': [
      'Product catalog with categories',
      'Weekly offers and promotions',
      'Delivery scheduling system'
    ],
    'Beauty Salon': [
      'Service menu with pricing',
      'Online appointment booking',
      'Stylist portfolio and reviews'
    ],
    'Cafe': [
      'Menu with high-quality images',
      'Online ordering system',
      'Special events and promotions'
    ],
    'Gym': [
      'Membership plans and pricing',
      'Class schedule and booking',
      'Trainer profiles and specialties'
    ],
    'Real Estate': [
      'Property listings with filters',
      'Virtual tours and image galleries',
      'Appointment scheduling for viewings'
    ]
  };
  
  const benefitsList = benefitMap[businessType] || [
    'Professional online presence',
    'Customer engagement features',
    'Business growth tools'
  ];
  
  benefitsList.forEach(benefit => {
    benefits += `• ${benefit}\n`;
  });
  
  return benefits;
}

function addBusinessSpecificBenefitsArabic(businessType, hasWebsite) {
  let benefits = '';
  
  const benefitMap = {
    'إلكترونيات': [
      'عرض المنتجات بشكل احترافي',
      'نظام حجز إلكتروني للمنتجات والخدمات',
      'المواصفات التقنية والمقارنات'
    ],
    'صيدلية': [
      'نظام حجز مواعيد إلكتروني',
      'كتالوج المنتجات مع خاصية البحث',
      'إدارة الوصفات الطبية'
    ],
    'سيارات': [
      'عرض السيارات بمواصفات مفصلة',
      'جدولة مواعيد الصيانة',
      'نظام حجز تجربة قيادة'
    ],
    'مطعم': [
      'قائمة طعام إلكترونية مع الصور',
      'نظام حجز الطاولات',
      'طلب أونلاين للتوصيل'
    ],
    'فندق': [
      'نظام حجز إلكتروني',
      'معرض الغرف والمرافق',
      'العروض الخاصة والباقات'
    ],
    'سوبرماركت': [
      'كتالوج المنتجات مع التصنيفات',
      'العروض الأسبوعية والترويجات',
      'نظام جدولة التوصيل'
    ],
    'صالون تجميل': [
      'قائمة الخدمات مع الأسعار',
      'حجز المواعيد أونلاين',
      'معرض أعمال المصممين وتقييماتهم'
    ],
    'مقهى': [
      'قائمة الطعام مع صور عالية الجودة',
      'نظام الطلب الإلكتروني',
      'الفعاليات الخاصة والترويجات'
    ],
    'نادي رياضي': [
      'باقات العضوية والأسعار',
      'جدول الحصص والحجز',
      'ملفات المدربين وتخصصاتهم'
    ],
    'عقارات': [
      'قوائم العقارات مع الفلاتر',
      'الجولات الافتراضية ومعرض الصور',
      'جدولة مواعيد المشاهدة'
    ]
  };
  
  const benefitsList = benefitMap[businessType] || [
    'حضور إلكتروني احترافي',
    'ميزات تفاعل العملاء',
    'أدوات نمو الأعمال'
  ];
  
  benefitsList.forEach(benefit => {
    benefits += `• ${benefit}\n`;
  });
  
  return benefits;
}

// PROGRESS CHECK FUNCTION (for triggers)
function checkMessageProgress() {
  const sheetNames = [
    'ELECTRONICS STORE_clean', 'PHARMACY_clean', 'CAR DEALER_clean',
    'REAL ESTATE AGENCY_clean', 'HOTEL_clean', 'SUPERMARKET_clean',
    'BEAUTY SALON_clean', 'CAFE_clean', 'GYM_clean', 'RESTAURANT_clean'
  ];
  
  let totalRows = 0;
  let processedRows = 0;
  
  for (const sheetName of sheetNames) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) continue;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;
    
    const messageRange = sheet.getRange(2, 15, lastRow - 1, 1);
    const messages = messageRange.getValues();
    
    let processed = 0;
    for (let i = 0; i < messages.length; i++) {
      if (messages[i][0] && messages[i][0] !== "") processed++;
    }
    
    totalRows += (lastRow - 1);
    processedRows += processed;
  }
  
  const percentage = totalRows > 0 ? Math.round((processedRows / totalRows) * 100) : 0;
  Logger.log(`📊 Message Generation Progress: ${processedRows}/${totalRows} (${percentage}%)`);
  
  return {
    processed: processedRows,
    total: totalRows,
    percentage: percentage
  };
}
