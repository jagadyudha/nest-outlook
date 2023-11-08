import { Body, Controller, Post } from '@nestjs/common';
import { SpreadsheetService } from 'src/services/spreadsheet.service';
import { OutlookService } from 'src/services/outlook.service';

@Controller('spreadsheet')
export class SpreadsheetController {
  constructor(
    private readonly outlookService: OutlookService,
    private readonly spreadSheetService: SpreadsheetService,
  ) {}

  @Post()
  async create(@Body() data: any) {
    const body = data.body ?? '';
    const subject: string = (data.subject ?? '').toLowerCase();
    if (subject.includes('penugasan audit')) {
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

  @Post('create_header')
  async createHeader() {
    const doc = await this.spreadSheetService.doc();
    const sheet = await doc.addSheet({
      title: new Date().toString(),
    });
    sheet.setHeaderRow(
      [
        'No',
        'ID Network',
        'Project Name',
        'Area',
        'Project Type',
        'SI',
        'Assignment Period',
        '',
        'Assigned Team',
        '',
        '',
        '',
        '',
        'Assignment Date',
        'Pemberitahuan',
        'Pelaporan',
        '',
      ],
      2,
    );
    sheet.setHeaderRow(
      [
        '',
        '',
        '',
        '',
        '',
        '',
        'Period',
        'Batch',
        'TL',
        'TM 1',
        'TM 2',
        'TM 3',
        'TM 4',
        '',
        '',
        'Target',
        'Actual',
      ],
      3,
    );
    const mergeHorizontal = [
      {
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      {
        startColumnIndex: 1,
        endColumnIndex: 2,
      },
      {
        startColumnIndex: 2,
        endColumnIndex: 3,
      },
      {
        startColumnIndex: 3,
        endColumnIndex: 4,
      },
      {
        startColumnIndex: 4,
        endColumnIndex: 5,
      },
      {
        startColumnIndex: 5,
        endColumnIndex: 6,
      },
      {
        startColumnIndex: 13,
        endColumnIndex: 14,
      },
      {
        startColumnIndex: 14,
        endColumnIndex: 15,
      },
    ];
    const mergetVertical = [
      {
        startRowIndex: 6,
        endRowIndex: 8,
      },
      {
        startRowIndex: 8,
        endRowIndex: 13,
      },
      {
        startRowIndex: 15,
        endRowIndex: 17,
      },
    ];
    mergeHorizontal.forEach(async (item) => {
      await sheet.mergeCells(
        {
          startRowIndex: 1,
          endRowIndex: 3,
          startColumnIndex: item.startColumnIndex,
          endColumnIndex: item.endColumnIndex,
        },
        'MERGE_ALL',
      );
    });
    mergetVertical.forEach(async (item) => {
      await sheet.mergeCells(
        {
          startRowIndex: 1,
          endRowIndex: 2,
          startColumnIndex: item.startRowIndex,
          endColumnIndex: item.endRowIndex,
        },
        'MERGE_ALL',
      );
    });
    return { message: 'success' };
  }
}
