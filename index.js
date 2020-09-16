// link to the spreadsheet: https://docs.google.com/spreadsheets/d/e/2PACX-1vSg28szyWFh5xWyvDTWo36QCm-4wlkJW3-6MMi-BDfk06eQbfMPbSTB4f1Q6QGKfA37IVcvDVAPq3Pu/pubhtml

const { GoogleSpreadsheet } = require("google-spreadsheet");

const credentials = {
  type: "service_account",
  project_id: "media-notas",
  private_key_id: "8187aRUuDa2zCjaTaYKMGKDGMBhd5QdfTr4927ff",
  private_key:
    "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCjbM5bPUYYMOHw\n/IUwtEEpqPPc+1iKIDRnAKwyVKNtFL4wr8Vxd1j0mDgrxbn6q00C1eCtatpRlL/1\nfeFcCE46ZeWfXHP4Zctojpm31RQN+6quKhm4LQd+ElYnADkPDdJGlQS1s5s8yYry\nLYBH0wfFVz2USw/QfWYKn57igjoZwpbTW0oC9PpR4OYS6o6wJHh8LIFCgXoPBvub\nLXZ4NBErpjyixLK0YV8N051WK22Yf4Y+NBOOdeWeHIsnwOoOddUtdwXqdP+zyJLA\nt7ZdmM4wf8y4V/Q02SLoBpSPj+32ZlNBoTwQfUkTp7QpUwFb2iEIMVvjQdvPhPGh\nxOnSPML5AgMBAAECggEAAIUuBn0UUiskqvxYzbIDM5dfuFw9+MmXtRy7z5i0oPok\nDVpf4+ez+ypZYm1JlWZSc0/8PD0W9xPExSqqix0VJ6svnxFfpqKnUjzC+UJ2wwEi\nNtX4OT1+dLyo9MyTwewit1oN1ui8laXUGQeDTSc7Mvn/Po+7HAgsKRw/OJwexHp4\nbAn3/zemby1Y1xTDui9PlpbjAG/hZgZHGL16jCjjzvEeYGKkEa7YAXbohY6ccSpw\nqdxhGNnxHrlIHZEAuMtS7wMo/MIKtqAfU3ykRKup4hCvS+BD+/uGjb9MHqLVNx2t\n4q2YYyEThVIrmNMIlgcUVB2po0QWavD8wweu6SXlwQKBgQDVD6L5qWDREfFwBFfu\nv5xvi9jxaXCUuCNRKxghNMbzfbAl9OjESY3KZ3P1w1bkiMQSXqfTsvQAnjWm2wwl\nqIxT5MN+zRIaWJvGPrHRE/3bsE22r/iWXcZ/yzU1378VKuotYSgmrc1gtC0Bxds3\nrig7emzFQLXDAmPBB2XYn+/bQQKBgQDEXFGntFsRr95CwtNIHKjUwbpfImAXTmPm\nivyXgmQa+L/cFMJfN9sr+DvHhStaPvdUNp4/RIdMVNg17v71bHOUMqbzWhdP6C8o\ntjaqdvACodtwm9pmR1ayHcR+03Wck8g5r2nRd5ekZWFV519bKoA5EAC2A10WxZOs\nNdFpclgRuQKBgQDNf852tXrPTFot69KQUbnmEOHHPya6Grzdrg4RASGfeqwqgAT+\nRd9/yLac5bLqrEtJWIjQ9HrKGc6vx/j4XZAz3qL8q3j5dluRI6lIetrQSUU7npDL\nH1m0quAAvXVFSmYiLOKYI+zCiCYc3qRpGQ5vB87flmF53NUwOh1uihzDwQKBgFeT\nSQO/x4Ia6sjhtXOK/K1u/Z0iarLaTmnrAP7ds6Hn4UHZrFrlQYXZv2eb+BrWzF4t\nweQ7vxAHIyrivaldxiqJcLZGLvF/f1Dr+3OJej/iSklt6TkGhh8IcbOSwfikXH+F\nwW8fpG04nfG/MGMrkGZiwb5rv5/BXLxIgG5EBg1ZAoGAcVbxFugOG5agJixfBC76\nZ9xlPKnBwIzDKU+Qdu0tI3acrhs5v6IYs7SGMUbBLw4bKFc21QAJocdQkKR8iDXX\nND8rv9qZEKIBiKse8a72Yo06yaJsWHS27PZZoAyZPvEfql1NkS7tlcd2jokx5JFi\nkTFcPKlQBguC3xrif+YXtes=\n-----END PRIVATE KEY-----\n",
  client_email: "media-notas@media-notas.iam.gserviceaccount.com",
  client_id: "111127082119173146972",
  auth_uri: "https://accounts.google.com/o/oauth2/auth",
  token_uri: "https://oauth2.googleapis.com/token",
  auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
  client_x509_cert_url:
    "https://www.googleapis.com/robot/v1/metadata/x509/media-notas%40media-notas.iam.gserviceaccount.com",
};

// in case of needing to change the spreadsheet for test purposes
const spreadsheetId = "1kO7L4EwHtvXf1WV3sTR3avBfqYZehqAnCtgHBPU7s8M";

const accessSheet = async () => {
  const doc = new GoogleSpreadsheet(spreadsheetId);
  await doc.useServiceAccountAuth(credentials);
  await doc.loadInfo();
  const sheet = doc.sheetsByIndex[0];
  await sheet.loadCells("A26");

  const config = {
    minGradeAverage: 70,
    classesAmount: (await sheet.getCell(25, 0)).value.split(" ")[5],
    studentGradeAverage: (p1, p2, p3) => (p1 + p2 + p3) / 3,
  };

  await sheet.getRows({ limit: 24 }).then((rows) => {
    rows.forEach((row) => {
      const gradeAverage = config.studentGradeAverage(
        parseInt(row.P1),
        parseInt(row.P2),
        parseInt(row.P3)
      );
      row.Nota_para_Aprovação_Final = 0;

      switch (true) {
        case row.Faltas > config.classesAmount * 0.25:
          row.Situação = "Reprovado por Falta";
          row.save();
          break;
        case gradeAverage >= 50 && gradeAverage < config.minGradeAverage:
          row.Nota_para_Aprovação_Final = (100 - gradeAverage).toFixed(0);
          row.Situação = "Exame Final";
          row.save();
          break;
        case gradeAverage < config.minGradeAverage:
          row.Situação = "Reprovado por Nota";
          row.save();
          break;
        case gradeAverage >= config.minGradeAverage:
          row.Situação = "Aprovado por Nota";
          row.save();
          break;
      }
      console.log(`
            ============================
            Aluno: ${row.Aluno}
            Matricula: ${row.Matricula}
            Situação: ${row.Situação}
            Nota para aprovação final: ${(
              row.Nota_para_Aprovação_Final / 10
            ).toFixed(0)}`);
    });
  });
};
accessSheet();
