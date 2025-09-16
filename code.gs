/**
 * ページを開いた時に最初に呼ばれるルートメソッド（門番）
 */
function doGet(e) {
  // 1. ログイン状況の取得
  const loggedInUser = PropertiesService.getUserProperties().getProperty('loggedInUser');
  
  if(loggedInUser) {
    // 2. 全てのマスターデータを取得し、勤怠画面の読み込み
    const masterData = getMasterData(); 
    const currentUser = masterData.instructors.find((row)=>{
      return String(row.id) === String(loggedInUser);
    });

    const scoreboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("運勢スコアボード");
    const scoreboardData = scoreboardSheet.getDataRange().getValues();

    let averageFortuneText = "-";
    let averageFortune = 0;
    for (let i = 1; i < scoreboardData.length; i++) {
      if (scoreboardData[i][0] == loggedInUser) {
        const totalPoints = scoreboardData[i][1] || 0;
        const totalCount = scoreboardData[i][2] || 0;
        if (totalCount > 0) {
          averageFortune = totalPoints / totalCount;
          averageFortune;
          if(averageFortune > 1) {
            averageFortuneText = "大吉";
          }else if(averageFortune >0) {
            averageFortuneText = "吉";
          }else{
            averageFortuneText = "凶";
          }
        }
        break;//一致する講師を見つけると処理終了
      }
    }

    if(currentUser&&currentUser.role === "塾長"){ 
      const template = HtmlService.createTemplateFromFile("view_management");
      template.masterData = masterData;
      template.loggedInUser = loggedInUser;
      return template.evaluate().setTitle("塾長専用画面");
    } else{
      const template = HtmlService.createTemplateFromFile("view_home");
      template.masterData = masterData;
      template.loggedInUser = loggedInUser;
      template.averageFortuneText = averageFortuneText;
      template.averageFortune = averageFortune.toFixed(2);
      return template.evaluate().setTitle("勤怠システム");
    }
  }else {
    const template = HtmlService.createTemplateFromFile('view_login');
    template.appUrl = ScriptApp.getService().getUrl();
    return template.evaluate().setTitle('ログイン');
  }
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * @ param{string} userId
 * @ param{string} password
 * @ param{boolen} 認証が成功したかどうか
 */

const authenticateUser = (userId, password) =>{
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('講師マスター');
  const data = masterSheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i ++){

    let row = data[i];
    console.log(`チェック中: シートのID(${typeof row[0]}): ${row[0]}, 入力されたID(${typeof userId}): ${userId}`);
    

    if(String(row[0]) == String(userId) && row[3] == password ){
      PropertiesService.getUserProperties().setProperty('loggedInUser', userId);
      return true;
    }
  }
  return false;
}


/**
 * 今日の出勤記録を見つけ、交通費を更新する
 * @param {string} fee 入力された新しい交通費
 */

function updateTransportationFee(fee){
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const data = logSheet.getDataRange().getValues();

  for(let i = data.length - 1; i >= 1; i--){
    const row = data[i];
    const actionType = row[1];

    if(actionType=="出勤"){
      logSheet.getRange(i+1, 6).setValue(fee);
      break;
    }
  }
  return "交通費が修正されました";
}

/**
 * マスターデータを取得する
 */
function getMasterData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // スプレッドシート全体を取得

  //名前でシートを取得
  var instructorSheet = ss.getSheetByName("講師マスター");
  var payRateSheet =ss.getSheetByName("コマ単価マスター");
  var levelSheet = ss.getSheetByName("授業レベルマスター"); 
  var scheduleSheet = ss.getSheetByName("時間割マスター");

  //オブジェクトに変換して、プロパティと値を対応させる

  // 「講師マスター」のデータを変換
  const instructorValues = instructorSheet.getRange(2, 1, instructorSheet.getLastRow() - 1, instructorSheet.getLastColumn()).getValues();
  const instructors = instructorValues.map(row => ({
    id: row[0],
    name: row[1],
    rank: row[2],
    role: row[4],
    transportationFee: row[5]
  }));

  //「コマ単価マスター」のデータを変換
  const payRateValues = payRateSheet.getRange(2, 1, payRateSheet.getLastRow() - 1, payRateSheet.getLastColumn()).getValues();
  const payRates = payRateValues.map(row => ({
    rank: row[0],
    level: row[1],
    pay: row[2]
  }));

  //「授業レベルマスター」のデータ変換
  const levelSheetValues = levelSheet.getRange(2, 1, levelSheet.getLastRow() - 1, levelSheet.getLastColumn()).getValues();
  const classLevels = levelSheetValues.map(row => ({
    level: row[0],
    basePay: row[1]
  }));


  // 「時間割マスター」のデータを変換
  const scheduleSheetValues = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow() - 1, scheduleSheet.getLastColumn()).getValues();
  const scheduleByPeriod = {}; 

  scheduleSheetValues.forEach((row)=>{
    const period = row[3];
    const classInfo = { name: row[0], startTime: Utilities.formatDate(new Date(row[1]), "JST", "HH:mm"), endTime: Utilities.formatDate(new Date(row[2]), "JST", "HH:mm") };
  
     if (!scheduleByPeriod[period]) {
      scheduleByPeriod[period] = [];
    }
    scheduleByPeriod[period].push(classInfo);

  })

  //全データを１つのオブジェクトにまとめる
  const masterData = {
    instructors: instructors,
    payRates: payRates,
    classLevels: classLevels,
    schedule: scheduleByPeriod
  };

  console.log(masterData.schedule);

  return masterData;
}


/**
 * 塾長専用画面の給与データを計算して返す
 * @param {string} targetInstructorId 対象の講師ID
 * @param {string} targetYearMonth 'YYYY-MM'形式の対象月 (例: '2025-08')
 * @return {Object} 計算された給与詳細データ
 */
function getSalaryDataForManagement(targetInstructorId, targetYearMonth) {
  const masterData = getMasterData();
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const logs = logSheet.getDataRange().getValues();

  // 対象年と月を数値として取得
  const targetYear = parseInt(targetYearMonth.split('-')[0]);
  const targetMonthNum = parseInt(targetYearMonth.split('-')[1]);

  // 計算結果を保存する変数
  let monthlyKomaSalary = 0;
  let monthlyTaskMinutes = 0;
  let monthlyTransportationFee = 0;
  let yearlyKomaSalary = 0;
  let yearlyTaskMinutes = 0;
  let yearlyTransportationFee = 0;
  
  const instructor = masterData.instructors.find(inst => inst.id == targetInstructorId);
  if (!instructor) { return { error: "講師が見つかりません。" }; }

  let yearlyTaskStartTime = null; //なぜ変数を分けるかチェック
  let monthlyTaskStartTime = null;

  // ログを1行ずつチェック
  logs.forEach(row => {
    const logInstructorId = row[0];
    // 対象講師のログでなければスキップ
    if (logInstructorId != targetInstructorId) { return; }

    const logDate = new Date(row[2]); // 日時はC列(インデックス2)
    const logYear = logDate.getFullYear();
    const logMonth = logDate.getMonth() + 1; // getMonthは0から始まるため+1
    const actionType = row[1]; // 種別はB列(インデックス1)
    const transportationFee = row[5] || instructor.transportationFee; // F列の交通費、なければマスターから


    // --- 年間累計の計算 ---
    if (logYear === targetYear && logMonth <= targetMonthNum) {
      if (actionType === '授業') {
        const classLevel = row[4]; // 授業レベルはE列(インデックス4)
        const payRate = masterData.payRates.find(rate => rate.rank === instructor.rank && rate.level === classLevel);
        if (payRate) {
          yearlyKomaSalary += payRate.pay;
        }
      } else if (actionType === '事務作業開始') {
        yearlyTaskStartTime = logDate;
      } else if (actionType === '事務作業終了' && yearlyTaskStartTime) {
        const diffMinutes = (logDate.getTime() - yearlyTaskStartTime.getTime()) / (1000 * 60); //ミリ秒で取得して分に変換
        yearlyTaskMinutes += diffMinutes;
        yearlyTaskStartTime = null;
      } else if(actionType === "出勤") {
        yearlyTransportationFee += transportationFee;
      }
    }

    // --- 月間給与の計算 ---
    if (logYear === targetYear && logMonth === targetMonthNum) {
      if (actionType === '授業') {
        const classLevel = row[4];
        const payRate = masterData.payRates.find(rate => rate.rank === instructor.rank && rate.level === classLevel);
        if (payRate) {
          monthlyKomaSalary += payRate.pay;
        }
      } else if (actionType === '事務作業開始') {
        monthlyTaskStartTime = logDate;
      } else if (actionType === '事務作業終了' &&  monthlyTaskStartTime) {
        const diffMinutes = (logDate.getTime() - monthlyTaskStartTime.getTime()) / (1000 * 60);
        monthlyTaskMinutes += diffMinutes;
        monthlyTaskStartTime = null;
      } else if(actionType === "出勤") {
        monthlyTransportationFee += transportationFee;
      }
    }
  });

  // 事務作業給を計算（時給1200円 = 分給20円）
  const monthlyTaskSalary = Math.ceil(monthlyTaskMinutes * 20); //小数点の切り上げ
  const yearlyTaskSalary = Math.ceil(yearlyTaskMinutes * 20);

  return {
    monthly: {
      komaSalary: monthlyKomaSalary,
      taskSalary: monthlyTaskSalary,
      transportationFee: monthlyTransportationFee,
      total: monthlyKomaSalary + monthlyTaskSalary + monthlyTransportationFee
    },
    yearly: {
      komaSalary: yearlyKomaSalary,
      taskSalary: yearlyTaskSalary,
      transportationFee: yearlyTransportationFee,
      total: yearlyKomaSalary + yearlyTaskSalary + yearlyTransportationFee
    }
  };
}
/**
 * 指定された月を「締め処理済み」として記録する
 * @param {string} yearMonth 'YYYY-MM'形式の対象月
 */
function closeMonth(targetYearMonth) {
  const closeMonthlyPayroll = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("締め管理");
  closeMonthlyPayroll.appendRow([targetYearMonth, '完了'])
  return '今月の締めが完了しました';
}



/**
 * 占いを行い、結果を運勢スコアボードに記録して返す
 * @param {string} instructorId 占いを引く講師のID
 * @return {Object} 占いの結果（メッセージとポイント）
 */
function getFortune(instructorId) {
  // --- 1. 必要なシートを取得する ---
  const fortuneSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("占いマスター");
  const scoreboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("運勢スコアボード");

  // --- 2. 占いマスターから、全おみくじのリストを取得 ---
  const fortunes = fortuneSheet.getRange("A2:C" + fortuneSheet.getLastRow()).getValues();

  // --- 3. ランダムに1つのおみくじを選ぶ ---
  const randomIndex = Math.floor(Math.random() * fortunes.length); //Math.random()で0〜1の間で数値を取り出し、列数をかけて切り下げ
  const selectedFortune = fortunes[randomIndex];
  
  const fortuneResult = {
    fortune: selectedFortune[0], 
    message: selectedFortune[1], // 2列目のメッセージ
    points: selectedFortune[2]   // 3列目のポイント
  };

  // --- 4. 運勢スコアボードを更新する ---
  const scoreboardData = scoreboardSheet.getDataRange().getValues();
  for (let i = 1; i < scoreboardData.length; i++) { // 1行目はヘッダーなので2行目から
    if (scoreboardData[i][0] == instructorId) { // A列の講師IDが一致したら
      const currentPoints = scoreboardData[i][1]  // B列の現在のポイント (空なら0)
      const currentCount = scoreboardData[i][2] ; // C列の現在の回数 (空なら0)
      
      // ポイントと回数を更新して、シートに書き込む
      scoreboardSheet.getRange(i + 1, 2).setValue(currentPoints + fortuneResult.points);
      scoreboardSheet.getRange(i + 1, 3).setValue(currentCount + 1);
      break; // 更新が終わったらループを抜ける
    }
  }

  // --- 5. 占いの結果を返す ---
  return fortuneResult;
}

/**
 * 毎月1日に、全講師の運勢スコアをリセットする関数
 */
function resetMonthlyFortunes() {
  // 1. 「運勢スコアボード」という名前のシートを操作する準備をします
  const scoreboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("運勢スコアボード");

  // 2. シートの名前を間違えて変更した場合の安全策
  if (scoreboardSheet) {
    // 3. データが入っている一番下の行が何行目か調べます
    const lastRow = scoreboardSheet.getLastRow();
    
    // 4. データに2行目以降がなければリセットされてしまうための安全策
    if (lastRow > 1) {
      // 5. B列の2行目からC列の最後まで（例: B2:C10）の範囲を選択し、中身を空っぽにします
      scoreboardSheet.getRange("B2:C" + lastRow).clearContent();
    }
  }
}

/**
 * 出勤打刻を記録する
 * @param {string} instructorId 選択された講師のID
 */
function recordClockIn(instructorId) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const timestamp = new Date(); // 現在時刻

  // 勤怠ログシートに新しい行を追加して書き込む
  logSheet.appendRow([instructorId,"出勤", timestamp]);

  return "出勤打刻を記録しました。"; // フロントエンドに返すメッセージ
}

/**
 * 退勤打刻を記録する
 * @param {string} instructorId 選択された講師のID
 */
function recordClockOut(instructorId) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const timestamp = new Date();

  // 打刻履歴シートに「退勤」の行を追加して書き込む
  logSheet.appendRow([instructorId, "退勤", timestamp]);

  return "退勤打刻を記録しました。";
}

/**
 *作業時間開始を記録する
 * @param {string} instructorId 選択された講師のID
 */

function recordTaskStart(instructorId) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打刻履歴');
  const timestamp = new Date();

  logSheet.appendRow([instructorId, '事務作業開始', timestamp]);
  return "事務作業開始を記録しました。";
}

function recordTaskEnd(instructorId){
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const timestamp = new Date();

  logSheet.appendRow([instructorId, "事務作業終了", timestamp]);
  return '事務作業終了を記録しました。';
}

/**
 * 休憩時間開始を記録する
 * @param {string} instructorId 選択された講師のID
 */

function recordBreakStart(instructorId) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打刻履歴');
  const timestamp = new Date();

  logSheet.appendRow([instructorId, "休憩開始", timestamp]);
  return '休憩開始を記録しました。';
}

/**
 * 休憩終了を記録する
 * @param {string} instructorId 選択された講師のID
 */

function recordBreakEnd(instructorId) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const timestamp = new Date();

  logSheet.appendRow([instructorId, "休憩終了",timestamp]);
  return "休憩終了を記録しました。";
}

/**
 * 担当した授業の報告と退勤打刻を記録する
 * @param {Array<Object>} taughtClasses 報告された授業情報の配列
 */
function recordTaughtClasses(taughtClasses, selectedPeriod) {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  const timestamp = new Date();
  const instructorId = PropertiesService.getUserProperties().getProperty('loggedInUser');

  // taughtClassesが空かどうかのチェック
  if (!instructorId) {
    // ログイン情報がない場合はエラーを返す
    throw new Error("ログイン情報が見つかりません。再ログインしてください。");
  }

  // 配列で受け取った授業を1行ずつ記録
  taughtClasses.forEach(classInfo => {
    logSheet.appendRow([
      instructorId,
      "授業",         // 種別
      timestamp,  //将来的に、授業ごとに押してもらうかも？
      classInfo.className, // 担当時限 (例: "1限")
      classInfo.level,      // 授業レベル (例: "中1")
      '',
      selectedPeriod
    ]);
  });
  
  // 通常の退勤打刻も記録する
  logSheet.appendRow([instructorId, "退勤", timestamp]);

  // 最後に占いを行い、その結果を返す
  const fortune = getFortune(instructorId);
  return fortune;
}

/**
 * ログイン状態をリセット（ログアウト）するための関数
 */
function logout() {
  PropertiesService.getUserProperties().deleteProperty('loggedInUser');
  console.log('ログアウトしました。通行証は無効になりました。');
}

