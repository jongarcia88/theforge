function applyTagRules() {
  const MAX_LOG_ENTRIES = 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName("Transactions");
  const ruleSheet = ss.getSheetByName("AutoTag");

  if (!txSheet || !ruleSheet) {
    throw new Error("Missing 'Transactions' or 'AutoTag' sheet.");
  }

  const txData = txSheet.getDataRange().getValues();
  const ruleData = ruleSheet.getDataRange().getValues();
  const headers = txData[0];

  const descCol = headers.indexOf("Description");
  const acctCol = headers.indexOf("Account");
  const instCol = headers.indexOf("Institution");
  const amtCol = headers.indexOf("Amount");
  const tagCol = headers.indexOf("Tags");

  if ([descCol, acctCol, instCol, amtCol, tagCol].includes(-1)) {
    throw new Error("Missing required columns in 'Transactions' sheet: Description, Account, Institution, Amount, Tags.");
  }

  let logCol = headers.indexOf("Tag Log");
  if (logCol === -1) {
    logCol = headers.length;
    txSheet.getRange(1, logCol + 1).setValue("Tag Log");
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  let totalTagsAdded = 0;

  for (let i = 1; i < txData.length; i++) {
    const tx = txData[i];
    const txDesc = (tx[descCol] || "").toString().toUpperCase();
    const txAcct = (tx[acctCol] || "").toString().toUpperCase();
    const txInst = (tx[instCol] || "").toString().toUpperCase();
    const txAmt = parseFloat(tx[amtCol]) || 0;
    let existingTags = (tx[tagCol] || "").toString().split(",").map(t => t.trim()).filter(Boolean);
    const existingLog = (tx[logCol] || "").toString();

    const tagsToAdd = new Set(existingTags);
    const matchLog = [];
    const startingCount = tagsToAdd.size;

    for (let j = 1; j < ruleData.length; j++) {
      const [tagString, descMatch, acctMatch, instMatch, minAmt, maxAmt, excludeList] = ruleData[j];
      if (!tagString) continue;

      const descOK = !descMatch || txDesc.includes(descMatch.toString().toUpperCase());
      const acctOK = !acctMatch || txAcct.includes(acctMatch.toString().toUpperCase());
      const instOK = !instMatch || txInst.includes(instMatch.toString().toUpperCase());
      const minOK = !minAmt || txAmt >= parseFloat(minAmt);
      const maxOK = !maxAmt || txAmt <= parseFloat(maxAmt);

      let excludeOK = true;
      if (excludeList) {
        const excludes = excludeList.toString().split(",").map(e => e.trim().toUpperCase());
        excludeOK = !excludes.some(word => txDesc.includes(word));
      }

      if (descOK && acctOK && instOK && minOK && maxOK && excludeOK) {
        const newTags = tagString.split(",").map(t => t.trim()).filter(Boolean);
        newTags.forEach(tag => {
          if (tag && !tagsToAdd.has(tag)) {
            tagsToAdd.add(tag);
            matchLog.push(`"${tag}" from rule: "${descMatch || '[any]'}" @ ${timestamp}`);
          }
        });
      }
    }

    const finalTags = Array.from(tagsToAdd).join(", ");
    if (finalTags !== existingTags.join(", ")) {
      txSheet.getRange(i + 1, tagCol + 1).setValue(finalTags);

      let updatedLogList = existingLog ? existingLog.split("\n") : [];
      updatedLogList = updatedLogList.concat(matchLog);
      if (updatedLogList.length > MAX_LOG_ENTRIES) {
        updatedLogList = updatedLogList.slice(-MAX_LOG_ENTRIES);
      }
      const updatedLog = updatedLogList.join("\n");

      txSheet.getRange(i + 1, logCol + 1).setValue(updatedLog);

      totalTagsAdded += tagsToAdd.size - startingCount;
    }
  }

  return `✅ Tags updated: ${totalTagsAdded} new tag(s) added.`;
}
