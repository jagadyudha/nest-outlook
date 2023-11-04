import { Injectable } from '@nestjs/common';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

@Injectable()
export class SpreadsheetService {
  async doc() {
    const serviceAccountAuth = new JWT({
      email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const doc = new GoogleSpreadsheet(
      '1u_SIGN3YuOVMZBZHTHPZzn_FM9WU-uRRhcTATQb0vdA',
      serviceAccountAuth,
    );
    await doc.loadInfo();
    return doc;
  }

  // async checkLastRow() {
  //   const doc = await this.doc();
  //   const sheet = doc.sheetsByIndex[0];
  //   const rows = await sheet.getRows();
  //   const lastRowId = rows.length ? rows[rows.length - 1].get('Id') : null;
  //   return lastRowId;
  // }

  async sendToExcel(data: {
    id: string;
    email: string;
    subject: string;
    body: string;
  }) {
    const doc = await this.doc();
    const sheet = doc.sheetsByIndex[0];
    await sheet.addRow(
      {
        Id: data.id,
        Email: data.email,
        Subject: data.subject,
        Body: data.body,
      },
      { insert: true },
    );
    return data;
  }
}
