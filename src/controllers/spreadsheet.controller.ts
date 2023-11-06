import { Body, Controller, Post } from '@nestjs/common';
// import { SpreadsheetService } from 'src/services/spreadsheet.service';
import { OutlookService } from 'src/services/outlook.service';

@Controller('spreadsheet')
export class SpreadsheetController {
  constructor(
    private readonly outlookService: OutlookService, // private readonly spreadSheetService: SpreadsheetService,
  ) {}

  @Post()
  async create(@Body() data: any) {
    const body = data.body ?? '';
    const header = [
      'ID',
      'ID NETWORK',
      'NETWORK NAME',
      'NETWORK TYPE',
      'AREA',
      'REGION',
      'MICROFINANCING',
      'REVIEW',
      'TL',
      'TM 1',
      'TM 2',
      'BATCH',
      'PERIOD',
    ];
    return this.outlookService.tableToJson({
      header,
      headerCount: 2,
      body,
    });
  }
}
