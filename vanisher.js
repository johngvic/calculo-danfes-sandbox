const fs = require("fs");

const retrieveRawData = (file) => {
  try {
    const data = fs.readFileSync(`./payloads/${file}_RAW.json`, "utf8");
    const rawData = JSON.parse(data);
    return rawData;
  } catch (err) {
    console.error(err);
  }
}

const retrievePendingData = (file) => {
  try {
    const data = fs.readFileSync(`./payloads/${file}_PENDING.json`, "utf8");
    const pendingData = JSON.parse(data);
    return pendingData;
  } catch (err) {
    console.error(err);
  }
}

const targetAlreadyExists = (matrix, target) => {
  // I have a target value { name, code, rps, nf, installment } and a matrix of values [ [{ name, code, rps, nf, installment }], ...].
  // If within this matrix I find any object containing `name`, `code`, `rps` and `nf` equal to the target, I return true. Otherwise, I return false.
  for (const row of matrix) {
    for (const obj of row) {
      if (obj.name === target.name && obj.code === target.code && obj.rps === target.rps && obj.nf === target.nf) {
        return true;
      }
    }
  }
  return false;
}

// customer C08047906000173 is a good delinquent example

(() => {
  const files = ["092025", "102025", "112025", "122025", "012026"];

  for (let index = 0; index < files.length; index++) {
    
    // testing purposes
    if (index == 1) {
      break;
    }

    const pendingData = retrievePendingData(files[index]);
    const eligibleRawFiles = files.slice(index);
    const register = [];

    for (const pending of pendingData) {

      // this will ignore a pending customer already vanished across all files and include only new customers
      if (targetAlreadyExists(register, pending)) {
        continue;
      }
      
      let companyInstances = [pending];

      for (const rawFile of eligibleRawFiles) {
        const raw = retrieveRawData(rawFile);
        const instances = raw.filter(({ name, code, rps, nf, installment }) => 
          name == pending.name &&
          code == pending.code &&
          rps == pending.rps &&
          nf == pending.nf &&
          installment != pending.installment
        );
        companyInstances = [...companyInstances, ...instances];
      }

      register.push(companyInstances);
    }

    fs.writeFileSync(`./payloads/${files[index]}_DELINQUENTS.json`, JSON.stringify(register, null, 2), () => {})
  }
})()