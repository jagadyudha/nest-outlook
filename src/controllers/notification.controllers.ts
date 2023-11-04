import { Body, Controller, Post } from '@nestjs/common';
// import { OutlookService } from 'src/services/outlook.service';
// import { firstValueFrom } from 'rxjs';
import { SpreadsheetService } from 'src/services/spreadsheet.service';

@Controller('notification')
export class NotificationController {
  constructor(
    // private readonly outlookService: OutlookService,
    private readonly spreadSheetService: SpreadsheetService,
  ) {}

  @Post('/sendEmailToExcel')
  async sendEmailToExcel(@Body() body: any) {
    // todo: accept json from pipedream
    return body;
  }
  // async sendEmailToExcel() {
  //   const messages: any = await firstValueFrom(
  //     this.outlookService.getUserMessages(),
  //   );
  //   const lastrowId = await this.spreadSheetService.checkLastRow();
  //   const filteredMessages = messages.value.filter((item) =>
  //     item.subject.includes('anu'),
  //   );
  //   const findIndex = filteredMessages.findIndex(
  //     (item) => item.id === lastrowId,
  //   );
  //   const index = findIndex > -1 ? findIndex : 0;
  //   const permitedData = filteredMessages.slice(
  //     index + 1,
  //     filteredMessages.length - 1,
  //   );

  //   permitedData.forEach((element) => {
  //     const { id, subject, body, sender } = element;
  //     this.spreadSheetService.sendToExcel({
  //       id,
  //       subject,
  //       email: sender.emailAddress.address,
  //       body: body.content,
  //     });
  //   });

  //   return permitedData;
  // }
}
