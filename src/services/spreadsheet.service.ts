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

  async sendToExcel(data: any) {
    const doc = await this.doc();
    const sheet = doc.sheetsByIndex[doc.sheetCount - 1];
    await sheet.addRow(data, { insert: true });
    return data;
  }

  async sendBulkToExcel(data: any, setting?: any) {
    const doc = await this.doc();
    let sheet = doc.sheetsByIndex[doc.sheetCount - 1];
    if (setting.sheetName) {
      sheet = doc.sheetsByTitle[setting.sheetName];
    }
    await sheet.addRows(data, { insert: true });
    return data;
  }
}
