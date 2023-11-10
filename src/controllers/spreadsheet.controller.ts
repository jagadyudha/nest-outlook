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
    const date = new Date(
      body.date ?? new Date().toUTCString(),
    ).toLocaleDateString('id-ID', {
      day: '2-digit',
      month: 'long',
      year: 'numeric',
      timeZone: 'Asia/Jakarta',
      timeZoneName: 'short',
    });
    const doc = await this.spreadSheetService.doc();
    const rows = await doc.sheetsByIndex[doc.sheetCount - 1].getRows({
      offset: 2,
    });
    if (subject.includes('penugasan audit')) {
      const table = this.outlookService.tableToJson({
        headerCount: 2,
        body,
      });
      const data = table.map((item, index) => {
        return [
          rows.length + index + 1,
          item['ID NETWORK'],
          item['NETWORK NAME'],
          item['AREA'],
          item['REVIEW'],
          '',
          item['PERIOD'],
          item['BATCH'],
          item['TL'],
          item['TM 1'],
          item['TM 2'],
          '',
          '',
          date,
        ];
      });
      return this.spreadSheetService.sendBulkToExcel(data);
    }

    if (subject.includes('pemberitahuan pemeriksaan')) {
      const table = [];
      const rowsObjects = rows.map((row) => {
        return { ...row.toObject(), index: row.rowNumber };
      });
      table.forEach(async (item) => {
        const find = rowsObjects.find(
          (row) => row['NETWORK ID'] === item['NETWORK ID'],
        );
        if (find) {
          rows[find.index].assign({});
          await rows[find.index].save();
        }
      });
    }
    return {
      message: 'Nothing.',
    };
  }

  @Post('create_header')
  async createHeader() {
    const doc = await this.spreadSheetService.doc();
    const sheet = await doc.addSheet({
      title: new Date()
        .toLocaleDateString('id-ID', {
          day: '2-digit',
          month: '2-digit',
          year: '2-digit',
        })
        .replace(/\//g, ' '),
    });
    sheet.setHeaderRow(
      [
        'No',
        'ID Network',
        'Network Name',
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
      1,
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
      2,
    );
    sheet.setHeaderRow(
      [
        'A',
        '',
        'B',
        '',
        'C',
        'D',
        'E',
        'F',
        'G',
        'H',
        'I',
        'J',
        'K',
        'L',
        'M',
        'N',
        'O',
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
          startRowIndex: 0,
          endRowIndex: 2,
          startColumnIndex: item.startColumnIndex,
          endColumnIndex: item.endColumnIndex,
        },
        'MERGE_ALL',
      );
    });
    mergetVertical.forEach(async (item) => {
      await sheet.mergeCells(
        {
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: item.startRowIndex,
          endColumnIndex: item.endRowIndex,
        },
        'MERGE_ALL',
      );
    });
    return { message: 'success' };
  }
}
