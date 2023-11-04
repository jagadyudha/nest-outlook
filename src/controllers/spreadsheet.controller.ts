import { Body, Controller, Post } from '@nestjs/common';
import { SpreadsheetService } from 'src/services/spreadsheet.service';

@Controller('spreadsheet')
export class SpreadsheetController {
  constructor(
    // private readonly outlookService: OutlookService,
    private readonly spreadSheetService: SpreadsheetService,
  ) {}

  @Post('/sendEmailToExcel')
  async sendEmailToExcel(@Body() body: any) {
    // todo: accept json from pipedream
    return body;
  }
}
