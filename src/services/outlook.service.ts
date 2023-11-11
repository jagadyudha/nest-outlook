import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { HttpService } from '@nestjs/axios';
import { Observable, mergeMap } from 'rxjs';
import { catchError, map } from 'rxjs';
import { AxiosResponse } from 'axios';
import { load as cheerioLoad } from 'cheerio';

@Injectable()
export class OutlookService {
  constructor(private readonly httpService: HttpService) {}

  refreshToken(): Observable<AxiosResponse<any>> {
    const body = {
      client_id: process.env.OUTLOOK_CLIENT_ID,
      scope: 'mail.read',
      refresh_token: process.env.OUTLOOK_TOKEN,
      grant_type: 'refresh_token',
      client_secret: process.env.OUTLOOK_CLIENT_SECRET,
    };
    return this.httpService
      .post(
        `https://login.microsoftonline.com/consumers/oauth2/v2.0/token`,
        body,
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        },
      )
      .pipe(
        map((response) => response.data),
        catchError((e) => {
          console.log(e.response.data);
          throw new HttpException(
            'Something went wrong',
            HttpStatus.INTERNAL_SERVER_ERROR,
          );
        }),
      );
  }

  getUserMessages(): Observable<AxiosResponse<any>> {
    return this.refreshToken().pipe(
      mergeMap((item: any) => {
        return this.httpService.get(
          `https://graph.microsoft.com/v1.0/me/messages?$select=sender,subject,body`,
          {
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
              Authorization: `Bearer ${item.access_token}`,
            },
          },
        );
      }),
      map((response) => response.data),
      catchError((e) => {
        console.log(e);
        throw new HttpException(
          'Something went wrong',
          HttpStatus.INTERNAL_SERVER_ERROR,
        );
      }),
    );
  }

  tableToJson(payload: { headerCount: number; body: string }) {
    const table = payload.body.match(/<table[^>]*>(.*?)<\/table>/g) ?? '';
    const $ = cheerioLoad(table.toLocaleString());
    const tableRows = $('table tr');
    const tableData = [];
    let rowIndex = 0;
    tableRows.each((_, row) => {
      let columnIndex = 0;
      $(row)
        .find('td, th')
        .each((_, cell) => {
          const rowspan = parseInt($(cell).attr('rowspan')) || 1;
          const colspan = parseInt($(cell).attr('colspan')) || 1;
          const value = $(cell).text() || '';
          for (let i = 0; i < rowspan; i++) {
            if (!tableData[rowIndex + i]) {
              tableData[rowIndex + i] = [];
            }
            while (tableData[rowIndex + i][columnIndex]) {
              columnIndex++;
            }
            for (let j = 0; j < colspan; j++) {
              tableData[rowIndex + i][columnIndex + j] = value;
            }
          }
          columnIndex += colspan;
        });
      rowIndex++;
    });
    const header = tableData.slice(0, payload.headerCount);
    const content = tableData.slice(payload.headerCount, tableData.length);
    const lastHeader = header[header.length - 1];
    const result = content.map((row) => {
      const obj = {};
      row.map((item, index) => {
        obj[lastHeader[index]] = item;
      });
      return obj;
    });
    return result;
  }
}
