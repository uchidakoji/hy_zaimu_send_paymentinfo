function sendPaymentHistory() {
    // readConfig 関数を呼び出して、設定シート上にある値を読み込む。返り値は連想配列。
    var configHash = readConfig();
    var answerSheet = configHash["回答シート名"];

    var sheetInput = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(answerSheet);
    var lastRowInput = sheetInput.getLastRow();
    var lastColInput = sheetInput.getLastColumn();
    var rg = sheetInput.getDataRange();
    
    for (var i = 1; i <= lastColInput; i++ ) {
        var colName = sheetInput.getRange(1, i).getValue(); // カラム名を取得  
        if (colName == "年会費納入方法") {var colNumPayment = i;}
        if (colName == "回生") {var colNumKaisei = i;}
        if (colName == "高校卒業時のルーム") {var colNumRoom = i;}
        if (colName == "氏名") {var colNumName = i;}
        if (colName == "氏名よみ") {var colNumNameKana = i;}
        if (colName == "郵便番号") {var colNumZipcode = i;}
        if (colName == "都道府県") {var colNumAdPref = i;}
        if (colName == "市町村番地") {var colNumberAdTown = i;}
        if (colName == "マンション名・部屋番号") {var colNumAdApartment = i;}
        if (colName == "メールアドレス") {var colNumEmail = i;}
        if (colName == "メール送信") {var colNumSentmail = i;}
    }
    
    var kaiseiInputOrg = String(sheetInput.getRange(lastRowInput, colNumKaisei).getValue()); //　入力された回生情報。数値から文字列に変換。文字列に変換しないと桁数を取得できない
    var roomInputOrg = String(sheetInput.getRange(lastRowInput, colNumRoom).getValue()); // 入力されたルーム情報。数値から文字列に変換。文字列に変換しないと桁数を取得できない
    var nameInputOrg = String(sheetInput.getRange(lastRowInput, colNumName).getValue()); // 入力された名前(漢字)
    var nameKanaInputOrg = String(sheetInput.getRange(lastRowInput, colNumNameKana).getValue()); // 入力された名前(かな)
    var sentmailInput = sheetInput.getRange(lastRowInput, colNumSentmail).getValue(); // メール送信状況。　処理済みかどうか判断するためのフラグとして使用
    var paymentMethod = sheetInput.getRange(lastRowInput, colNumPayment).getValue();　// 希望する支払い方法
    var mailAddress = sheetInput.getRange(lastRowInput, colNumEmail).getValue();　// Eメールアドレス

    if (sentmailInput == "yes"){
        return; //メール送信欄に"yes"と入力されている場合は処理済みなのでここでScript終了
    } 

    if(kaiseiInputOrg.length === 1){
        var kaiseiInput = "0" + kaiseiInputOrg; //回生が1桁の場合は、最初に0を足す
    } 
    else{
        var kaiseiInput = kaiseiInputOrg;
    }
    
    var roomInput = roomInputOrg.slice(1); // 入力されたルーム情報の下一桁を取得
    var nameInput = remove_space(nameInputOrg);　// 姓名の間にスペースを入力する人もいるので、定義したremove_space 関数を呼び出して、全角・半角スペースを取り除く
    var nameKanaInput = remove_space(nameKanaInputOrg);
    var nameKanaInput = convert_small_to_large(nameKanaInput); // 定義したconvert_small_to_large関数を呼び出して、拗音（ようおん）「ゃゅょ」や促音（そくおん）「っ」を通常のひらがなに変換。データベース内で一部の会員かな氏名の拗音や促音が通常のひらがなで入力されているため。

    Logger.log("nameInput: " + nameInput);
    Logger.log("roomInput: " + roomInput);

    var thisFiscalYearInt = check_fiscalyear();　//　定義したcheck_fiscalyear関数を呼び出して、現在の会計年度を取得
    var thisFiscalYear = String(thisFiscalYearInt) + "年度"
    //var thisFiscalYear = "2018年度";
    Logger.log("thisFiscalYear: " + thisFiscalYear);
    var currentQuarter = check_currentQuarter(thisFiscalYear);　//　定義したcheck_current_quarter関数を呼び出して、現在がどの四半期なのかを取得
    Logger.log(currentQuarter);

    var nenkaihiDB = configHash["年会費支払履歴DB"];
    var nenkaihiSheet = configHash["年会費支払履歴シート名"];
    
    var sheetPayRec = SpreadsheetApp.openById(nenkaihiDB).getSheetByName(nenkaihiSheet); // 「年会費支払履歴データベース_Prod」スプレッドシートの「年会費支払履歴」シートを取得。
    var lastRowPayRec = sheetPayRec.getLastRow();
    var lastColPayRec = sheetPayRec.getLastColumn();
    var values = sheetPayRec.getSheetValues(1, 1, lastRowPayRec, lastColPayRec); //　年会費支払履歴シートの全てのセルを選択し、配列化する

    for (var i = 0; i<= lastColPayRec - 1; i++) { // 年会費支払履歴シート内で、今年の会計年度の情報がどの列にあるのか確認.
        if(values[0][i] === thisFiscalYear){
        var colNumThisFiscalYear = i; 
        break;    
        }
    }

    // 入力された情報を基に会員の情報が年会費支払履歴シート内にあるかを確認。定義したsearch_member関数を使用する。戻り値は配列
    var arrayMatch = search_member(nameInput, nameKanaInput, kaiseiInput, roomInput, values);
    Logger.log("検索結果 戻り値:");
    Logger.log(arrayMatch);

    var groupName = configHash["グループ名"];
    
    // 検索結果によって異なる処理
    if (arrayMatch.length === 0) { // 検索結果が0件
        
        var matchResult = "notFound";
        Logger.log("matchResult: " + matchResult);
        
        var message = "大変申し訳ありませんが、" + nameInputOrg +　" 様の回生とお名前をデータベースから検索することができませんでした。" + groupName + "事務局で確認を行い、入力されたこちらのメールアドレスに連絡いたします。"
        + "<br>念のため入力された回生、高校卒業時のルーム、氏名、氏名よみに間違いがなかったかをご確認ください。";
        
    }
    else if (arrayMatch.length > 1) { // 検索結果が複数
        
        var matchResult = "duplicated";
        Logger.log("matchResult: " + matchResult);
        
        var message = nameInputOrg +　" 様の回生とお名前に一致する検索結果が" + arrayMatch.length + "件あります。どちらが" + nameInputOrg + " 様の会員情報かこちらのメールに返信する形で教えていただけますでしょうか？。<br>"
        + groupName + "事務局で確認を行い、改めて連絡させていただきます。";
        
        // 検索結果を表示
        var duplicatedResults = "<br><br>----------------------------";
        
        for (var i = 0; i<= arrayMatch.length - 1; i++) {
        //Logger.log(arrayMatch[i]);
        duplicatedResults = duplicatedResults + "<br>会員番号: <b>" + arrayMatch[i][1] + "</b>"
        + "<br>卒業時のルーム: <b>6" + arrayMatch[i][5] + "ルーム</b>"
        + "<br>氏名: <b>" + arrayMatch[i][2] + "</b><br>"
        + "----------------------------";  
        }
        
        message = message + duplicatedResults;
        
    }
    else if (arrayMatch.length === 1) {　//　検索結果が1件. 処理を進める
        
        var matchResult = "found";
        Logger.log("matchResult: " + matchResult);
        
        var matchRow = arrayMatch[0][0];
        var memberNumber　= arrayMatch[0][1];
        
        //　今年度を含めた3年間の支払い履歴を調べて、テーブルで表示。未納年数もカウント
        var countNopay = 0;　//　未納年数
        var table = "<table border=1　style= \"font-family:helvetica,arial,meiryo,sans-serif;font-size:10.5pt\"><tr style=\"background:#ccccff\"><th>年度</th><th>納入状況</th></tr>";
        
        //卒業から2年以上経過していない場合は、過去3年の履歴を調べる必要がない。
        //"卒業年度の次年度 - 回生 = 1959"であるため、その値を利用して確認する
        var countSinceGrad = thisFiscalYearInt - sheetInput.getRange(lastRowInput, colNumKaisei).getValue() - 1959;
        var thisYearPayrec = "";
        var oneYearAgoPayrec = "";
        var twoYearAgoPayrec = "";

        for (var i = 0; i <= 2; i++) {
        var payrec = convert_payrec(values[matchRow][colNumThisFiscalYear - i], groupName); // 定義したconvert_payrec関数でシート内の 1,B といった記号を実際の振り込み情報に変換
        var year = values[0][colNumThisFiscalYear - i];
        
            if (i === 0){ // 今年度の納入情報
                thisYearPayrec = payrec;
            }
            else if (i === 1){ // 1年前の納入情報
                if(i > countSinceGrad){
                    payrec = "駒東在校中のため不要";
                }
                oneYearAgoPayrec = payrec;
            }
            else if (i === 2){ //　2年前の納入情報
                if(i > countSinceGrad){
                    payrec = "駒東在校中のため不要";
                }
                twoYearAgoPayrec = payrec;
            }
            // 納入履歴の情報を基にテーブルを作成
            if (payrec.indexOf("納入済") >= 0){
                var table = table + "<tr><td>" + year + "</td><td>" + payrec + "</td></tr>";    
            }
            else if (payrec.indexOf("未納入") >= 0){      
                var table = table + "<tr><td>" + year + "</td><td style=\"background:#ffcccc\">" + payrec + "</td></tr>";
                countNopay++;      
            }
            else if (payrec.indexOf("駒東在校中") >= 0){
                var table = table + "<tr><td>" + year + "</td><td>" + payrec + "</td></tr>"; 
            }
        }
        
        table = table + "</table>"
        
        Logger.log("thisYearPayrec: " + thisYearPayrec);
        Logger.log("oneYearAgoPayrec: " + oneYearAgoPayrec);
        Logger.log("twoYearAgoPayrec: " + twoYearAgoPayrec);
        Logger.log("countNopay: " + countNopay);
        Logger.log(table);
        
        // ここより　各種条件によって、会員への異なるメッセージを作成.定義したcompose_message関数を使用する。戻り値はString
        var message = compose_message(paymentMethod, thisYearPayrec, oneYearAgoPayrec, countNopay, configHash);
        Logger.log(message);
    }

    // Eメールレポートを送るためにメッセージを加工
    var bankaccount = configHash["銀行名"];
    var address = "<p>〒" + sheetInput.getRange(lastRowInput, colNumZipcode).getValue() + "<br>"
    + sheetInput.getRange(lastRowInput, colNumAdPref).getValue() + " " + sheetInput.getRange(lastRowInput, colNumberAdTown).getValue() + " " + sheetInput.getRange(lastRowInput, colNumAdApartment).getValue() + "</p>";
    
    // 自動引き落とし希望者へのメッセージを作成する。申し込み月によって、内容を変更。１，２，３月に申し込んだ場合、状況によって間に合わないため
    var withdrawalInfo = "入力された以下のご住所に自動引き落とし申し込み用紙を送付いたします。<br>ご記入の上同封の封筒にてご返信ください。" + address
    var thisYear = Number(Utilities.formatDate(new Date(), 'JST', 'yyyy')); // 年を取得
    var thisMonth = Number(Utilities.formatDate(new Date(), 'JST', 'MM')); // 月を取得
    if (thisMonth < 3) {
        var nextYear = String(thisYear);
        withdrawalInfo = withdrawalInfo + "年会費の口座引落としは毎年4月27日を予定しており、" + nextYear + "年度分からの取り扱いになります。金融機関との登録手続きには約40日を要します。お早目にお申込み頂けますようお願い申し上げます。";
    }
    else if (thisMonth === 3){
        var nextYear = String(thisYear);
        withdrawalInfo = withdrawalInfo + "年会費の口座引落としは毎年4月27日を予定しております。金融機関との手続きに約40日かかるため、書類の確認に時間を要した場合など、" + nextYear + "年度からのお取り扱いに間に合わない場合があることを予めご了承願います。";
    }
    else{
        var nextYear = String(thisYear + 1);
        withdrawalInfo = withdrawalInfo + "年会費の口座引落としは毎年4月27日を予定しており、" + nextYear + "年度分からの取り扱いになります。"
    }

    var nameKatakana = nameKanaInput.replace(/[\u3041-\u3096]/g, function(s) { // 銀行振込ガイダンス用にひらがな氏名をカタカナに変換 
            return String.fromCharCode(s.charCodeAt(0) + 0x60);
        } ); 

    // compose_message() 関数からの戻り値に含まれている"$変数”を実際のテキストに置換する
    message = message.replace("$thisFiscalYear",thisFiscalYear);
    message = message.replace("$nameInputOrg", nameInputOrg);
    message = message.replace("$withdrawalInfo", withdrawalInfo);
    message = message.replace("$bankaccount", bankaccount);
    message = message.replace("$memberNumber", memberNumber);
    message = message.replace("$nameKatakana", nameKatakana);
    message = "<p>" + message + "</p>";
    Logger.log("message: " + message);

    var tableTitle = "<p><b>" + nameInputOrg + " (会員番号:" + memberNumber + ") 様の直近3年間の年会費納入状況:</b>";
    var tableNotes = "<b>注1：</b>" + groupName + "の会計年度は4月1日～3月31日で、現在は" + currentQuarter + "です。"
    + "<br><b>注2：</b>年会費の納入状況は定期的にデータベースが更新されていますが、必ずしも最新情報ではないことをご了承ください。</p>";

    var userInputs = "<p style= \"font-family:helvetica,arial,meiryo,sans-serif;font-size:9pt\">### 以下、入力された情報 ###"
    + "<br>[回生]: " + kaiseiInputOrg
    + "<br>[高校卒業時のルーム]: " + roomInputOrg
    + "<br>[氏名]: " + nameInputOrg
    + "<br>[氏名よみ]: " + nameKanaInputOrg
    + "<br>[年会費納入方法]: " + paymentMethod
    + "</p>"

    var mailFrom = configHash["送信者"];
    var replyToAddress = configHash["返信先"];
    var bccAddress = configHash["bccアドレス"];

    if (matchResult == "found") {　// 年会費納入履歴が検索できた場合のメッセージ  
        var header = "<p>" + kaiseiInputOrg + "回生 " + nameInputOrg + " 様 (会員番号: " + memberNumber + ")<br><br>" + groupName + "年会費納入に関するお問い合わせありがとうございました。</p>";
        var footer = "<p>この情報を基に" + groupName + "より電話やEメール（このEメールを除く）にて年会費支払いの督促を行うことはございません。不審な督促を受けた場合には、" + groupName + "までご連絡ください。"
        + "<br><br>何かご質問がございましたら、このメールにご返信ください。"
        + "<br><br>" + mailFrom + "</p>";
        var mailBody = header + message + tableTitle + table + tableNotes + footer + userInputs;  
    }
    else {　// 年会費納入履歴が検索できなかった場合のメッセージ  
        var header = kaiseiInputOrg + "回生 " + nameInputOrg + " 様 <br><br>" + groupName + "年会費納入に関するお問い合わせありがとうございました。";    
        var footer = "<p>何かご質問がございましたら、このメールにご返信ください。"
        + "<br><br>" + mailFrom + "</p>";    
        var mailBody = header + message + footer + userInputs;    
    }

    mailBody = "<body style= \"font-family:helvetica,arial,meiryo,sans-serif;font-size:10.5pt\">" + mailBody + "</body>";
    Logger.log("mailBody: " + mailBody);

    //Mail Subject を作成。テスト環境の場合は＜QA Test>をSubjectに入れる   
    var mailSubject = kaiseiInputOrg + "回生 " + nameInputOrg + " 様　" + groupName + " 年会費納入に関するお問い合わせありがとうございました。";
    if(configHash["環境"] === "qa"){
        mailSubject = "<QA TEST> " + mailSubject;
    }

    if (message.indexOf("thebase.in/payments") >= 0) { //クレジットカード支払いの案内をPDFで送るかどうかを、支払いサイトのURLがメール内にあるかどうかで判断。*** 注 *** 2018年4月現在、クレジットカードによる支払いは導入していない    
        var creditPDF = configHash["クレジットPDFファイル"]; //PDFファイルのFile IDを取得
        var file = DriveApp.getFileById(creditPDF);　
        
        MailApp.sendEmail({      
        to: mailAddress,
        subject: mailSubject,
        htmlBody: mailBody,
        replyTo: replyToAddress,
        bcc: bccAddress,
        name: mailFrom,
        attachments: [file.getAs(MimeType.PDF)]      
        });  
        
    }
    else {
        MailApp.sendEmail({      
        to: mailAddress,
        subject: mailSubject,
        htmlBody: mailBody,
        replyTo: replyToAddress,
        name: mailFrom,
        bcc: bccAddress      
        });   
    }

    sheetInput.getRange(lastRowInput, colNumSentmail).setValue("yes"); // メール送信後　”メール送信”コラムにyesを記入

}

//##################################################
//##################################################
function compose_message(paymentMethod, thisYearPayrec, oneYearAgoPayrec, countNopay, configHash) {

// $変数　は　繰り返しとなるので、後で実際のテキストに置換
var inputFormURL = configHash["入力フォームURL"];
var groupName = configHash["グループ名"];

if (paymentMethod.indexOf("年会費納入状況") >= 0){ //過去3年間の支払い状況のみを確認する場合 
var message = "$nameInputOrg様の直近3年間の年会費納入状況は以下の通りです。"
    if (thisYearPayrec.indexOf("未納入") >= 0){   
        message = message + "今年度分($thisFiscalYear)の年会費納入状況は<b><u>未納入</u></b>となっております。"
        + "<br><br>年会費の支払い方法についての詳細が知りたい場合は、再度<A href=\"" + inputFormURL + "\">こちらのリンク</A>より必要な情報の入力をお願いいたします。"   
    }
    else {   
        message = message + "すでに今年度分($thisFiscalYear)の年会費を頂いております。大変ありがとうございました。"   
    }
}
else if (thisYearPayrec == "納入済：銀行口座からの自動引き落とし"){ //銀行口座引き落としですでに今年度分年会費を支払い済みのケース
 
    var message = "$nameInputOrg様からは「銀行口座からの自動引き落とし」にてすでに今年度分($thisFiscalYear)の年会費を頂いております。大変ありがとうございました。"
    if(paymentMethod.indexOf("変更") >= 0){ // 引き落とし口座の変更希望の会員に対するコメント     
        message = message + "<br><br>自動引き落とし口座の変更を希望されていますので、$withdrawalInfo";     
    }
 
}
else if (thisYearPayrec == "納入済：" + groupName + "への振込"){ //邦友会口座への振込ですでに今年度分年会費を支払い済みのケース
    var message = "$nameInputOrg様からはすでに今年度分($thisFiscalYear)の年会費を頂いております。大変ありがとうございました。"
    
    if(paymentMethod.indexOf("自動引き落とし") >= 0){ // 今年度分納入　かつ　自動引き落とし希望の会員に対するコメント    
        message = message + "<br><br>$withdrawalInfo";       
    }
  
}
else if (thisYearPayrec.indexOf("未納入") >= 0){ //今年度分年会費が未納入のケース
    var thisMonth = Number(Utilities.formatDate(new Date(), 'JST', 'MM')); // 月を取得
    Logger.log("thisMonth: " + thisMonth);
    //thisMonth = 5;
 
    /* ここに　前年度未納入で 引き落とし口座変更の希望を出している人のメッセージを送るプログラムをかく　*/

    if ( (oneYearAgoPayrec == "納入済：銀行口座からの自動引き落とし")　&& ( (thisMonth >= 4) && (thisMonth <= 6) ) ){　//　第一四半期はまだ自動引き落としの結果がDBに反映されていないことを考慮
        var message = "$nameInputOrg様は「銀行口座からの自動引き落とし」を利用して、" + groupName + "年会費の振込みをして頂いております。"
        + "<br>今年度分の支払い状況は未納入となっておりますが、これは毎年4月27日に行われている振込の結果がまだデータベースに反映されていない可能性がございます。"
        + "<br>お手数ではございますが、7月頃にまた改めて年会費支払いに関しての問い合わせをしていただけますでしょうか。よろしくお願いいたします。";   
    }
    else {
        var message = "$nameInputOrg様の今年度分($thisFiscalYear)の年会費納入状況は<b><u>未納入</u></b>となっております。"
        var nopayAmount = 2000 * countNopay;
        Logger.log("nopayAmount: " + nopayAmount);
    
        if(countNopay > 1){
            var ruleURL = configHash["会則URL"];
            var message = message + "<br>また、今年度を含めた過去3年間で<b><u>" + countNopay + "年分未納入</u></b>となっております。"
            var message = message + "<br><br><A href=\"" + ruleURL + "\">" + groupName + " 会計規約第6条</A>により「過去の年会費未納分については最大3年に亘り遡り請求することが出来る。」となっており、過去の未納分を含めた<b><u>" + nopayAmount + "円</u></b>をお支払いいただけると大変助かります。";       
        }
        else if (countNopay == 1){
            var message = message + "<br>今年度分の年会費<b>" + nopayAmount + "円</b>をお支払いいただけると助かります。"       
        }
        
        //// ここより支払い方法の説明
        if(paymentMethod.indexOf("自動引き落とし") >= 0){
        message = message + "<br><br>$withdrawalInfo";
        
            if(countNopay > 0){
                message = message + "<br><br>また、お手数ではございますが<b><u>未納入分" + nopayAmount + "円</u></b>に関しましては、以下の銀行口座にお振込みをお願いいたします。<br>振込依頼人名は「$memberNumber$nameKatakana」とご入力下さい。$bankaccount";
            }
        
        }
        else if (paymentMethod.indexOf("銀行振込") >= 0){
            message = message + "<br><br>お手数ではございますが、以下の銀行口座にお振り込みをお願いいたします。<br>振込依頼人名は「$memberNumber$nameKatakana」とご入力下さい。$bankaccount";          
        }
        else if (paymentMethod.indexOf("クレジットカード") >= 0){
            var creditURL = configHash["クレジットカードURL"];
            message = message + "<br><br><A href=\"" + creditURL + "\">こちらのリンク</A>よりクレジットカードにてお支払いが可能です。添付のPDFファイル内の手順に従ってお支払いいただきますようお願いいたします。"
            + "<br><font size=\"2\"><b>注：</b><br> - 当決済サービスを利用いただけるのは、VISA、Master、AMEX、JCBカードのみです。"
            + "<br> - 支払い方法は1回払いのみです"
            + "<br> - 当決済サービスは<A href=\"https://thebase.in/payments\">BASE株式会社の決済サービス</A>を利用しており、カード番号・個人情報等は暗号化され保護されています。</font>";              
        }
    
    }
 
 }

return message;

}

//##################################################
//##################################################
function remove_space(name) {
 
    var targetStrHanSpace = " "; // まずは半角スペースを取り除く
    var regExp = new RegExp(targetStrHanSpace, "g");
    var name = name.replace(regExp, "");
        
    var targetStrZenspace = "　"; // 次に全角スペースを取り除く
    var regExp = new RegExp(targetStrZenspace, "g");
    var name = name.replace(regExp, "");
    
    return name;
  
}
//##################################################
//##################################################
function convert_payrec(payrec, groupName){
 
    payrec = String(payrec) //値を文字列に変換
  
    if(payrec.indexOf("B") >= 0){   
        var statusPayrec = "納入済：銀行口座からの自動引き落とし";    
    }
    else if (payrec.indexOf("1") >= 0){  
        var statusPayrec = "納入済：" + groupName + "への振込";
        if ((payrec.indexOf("T") >= 0) || (payrec.indexOf("Z") >= 0)){　//記号T　および　Z　が入力されている場合の処理     
            var statusPayrec = statusPayrec + "<br>ただし、銀行口座からの自動引き落としは失敗";      
        }
    }
    else {   
        var statusPayrec = "未納入";
        if ((payrec.indexOf("T") >= 0) || (payrec.indexOf("Z") >= 0)){　//記号T　および　Z　が入力されている場合の処理     
            var statusPayrec = statusPayrec + ":銀行口座からの自動引き落としの失敗";      
        }
    }
 
    return statusPayrec;
  
}
//##################################################
//##################################################
function check_fiscalyear(){
  
    var thisYear = Number(Utilities.formatDate(new Date(), 'JST', 'yyyy'));
    var thisMonth = Number(Utilities.formatDate(new Date(), 'JST', 'MM'));
    if(thisMonth <= 3){
        thisYear = thisYear - 1
    }

    /*if(thisMonth <= 3){
        thisYear = String(thisYear - 1) + "年度";    
    }
    else {   
        thisYear = String(thisYear) + "年度";    
    }*/
  
    return thisYear;
  
}
//##################################################
//##################################################
function check_currentQuarter(year){

    var thisMonth = Number(Utilities.formatDate(new Date(), 'JST', 'MM'));
    //thisMonth = 5;
    if (thisMonth <= 3){
        var q = "第4四半期";    
    }
    else if (thisMonth <= 6){  
        var q = "第1四半期";    
    }
    else if (thisMonth <=9){  
        var q = "第2四半期";  
    }
    else if (thisMonth <=12){  
        var q = "第3四半期";  
    }

    var currentQuarter = year + q;
    
    return currentQuarter;

}
//##################################################
//##################################################
function convert_small_to_large(name){

    //　拗音や促音を大きい文字に変換
    name = name.replace(/[ぁぃぅぇぉゕゖっゃゅょゎ]/g, function(s) {
            return String.fromCharCode(s.charCodeAt(0) + 1);
        });

    return name;

}
//##################################################
//##################################################
function search_member(name, nameKana, kaisei, room, values) {

    Logger.log("Room: " + room);

    var returnArray = [];

    for (var i = 1; i <= values.length-1; i++){
    
        var memberNumber = String(values[i][0]); // 年会費支払履歴シートの会員番号を取得。文字列化しておく
        if (memberNumber.length == 4) {memberNumber = "0" + memberNumber} // 会員番号が4ケタだった場合、最初に0を足す
        
        var kaiseiPayRec = memberNumber.substr(0, 2);　//　年会費支払履歴シート内の会員番号の最初の2ケタ(回生情報)を取得
        var roomPayRec = memberNumber.substr(2, 1); // 会員番号から卒業時ルーム下1桁を取得
        var namePayRec = values[i][1]; //　年会費支払履歴シート内の漢字氏名を取得　
        var nameKanaPayRec = values[i][2];　//　年会費支払履歴シート内のかな氏名を取得
        var memberNumber = values[i][0]; // 年会費支払履歴シートの会員番号を取得
        
        if ( ( (name == namePayRec) || (nameKana == nameKanaPayRec) ) && (kaisei == kaiseiPayRec) ) { // まずは　回生＋(漢字 or よみ)　の一致で検索      
            returnArray.push([i, memberNumber, namePayRec, nameKanaPayRec, kaiseiPayRec, roomPayRec]);      
        }
        
    }
    
    Logger.log("検索結果　1: ");
    Logger.log(returnArray);
    
    if (returnArray.length <= 1) { // 検索結果が0または1つのレコードしかもたなければ、結果を戻す    
        return returnArray;    
    }
    else if (returnArray.length > 1) {　// 検索結果が2つ以上のレコードを持つ場合は、更なる確認を行う    
        var returnArray2 = [];
        
        if( room != "" ){ //回生の情報が入力されている場合は、それを利用する
        
            for (var j = 0; j <= returnArray.length - 1; j++){
                
                if (returnArray[j][5] == room){          
                    returnArray2.push( returnArray[j] );          
                }
                
            }
        
        }
        
        if( returnArray2.length === 1 ){return returnArray2;}
        
        for (var j = 0; j <= returnArray.length - 1; j++){
        
            if (returnArray[j][2] === name){ // 漢字氏名が同じで、検索結果が複数表示されている場合
            
                if ( returnArray[j][3] === nameKana ) {          
                    returnArray2.push( returnArray[j] );
                }      
            
            }
            else if (returnArray[j][3] === nameKana){ // かな氏名が同じで、検索結果が複数表示されている場合
            
                if ( returnArray[j][2] === name ) {          
                    returnArray2.push( returnArray[j] );
                }      
            
            }
        
        }
        Logger.log("検索結果　2: ");
        Logger.log(returnArray2);
        return returnArray2;    
    }

}
////////////////////
