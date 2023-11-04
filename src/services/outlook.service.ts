import { HttpException, HttpStatus, Injectable } from '@nestjs/common';
import { HttpService } from '@nestjs/axios';
import { Observable, mergeMap } from 'rxjs';
import { catchError, map } from 'rxjs';
import { AxiosResponse } from 'axios';

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
}
