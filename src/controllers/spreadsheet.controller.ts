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
    const date = data.date ? new Date(data.date) : new Date();
    const doc = await this.spreadSheetService.doc();
    const stringMonth =
      'Bulan ' +
      date.toLocaleDateString('id-ID', {
        month: 'long',
      });
    const sheet = doc.sheetsByTitle[stringMonth];
    if (!sheet) {
      return { message: 'Sheet not found' };
    }
    const rows = await doc.sheetsByTitle[stringMonth].getRows({
      offset: 2,
    });
    if (subject.includes('penugasan audit')) {
      const table = this.outlookService.tableToJson({
        headerCount: 2,
        body,
      });
      const currentDate = date
        .toLocaleDateString('id-ID', {
          weekday: 'long',
          day: '2-digit',
          month: 'long',
          year: 'numeric',
          timeZone: 'Asia/Jakarta',
        })
        .toString();
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
          ` ${currentDate}`,
        ];
      });
      return this.spreadSheetService.sendBulkToExcel(data, {
        sheetName: stringMonth,
      });
    }

    if (subject.includes('pemberitahuan pemeriksaan audit')) {
      const OFFSET = 3;
      await doc.sheetsByTitle[stringMonth].loadHeaderRow(OFFSET);
      const table = this.outlookService.tableToJson({
        headerCount: 1,
        body,
      });
      const rowsObjects = rows.map((row) => {
        return { ...row.toObject(), index: row.rowNumber - OFFSET - 1 };
      });
      table.forEach(async (item) => {
        const find = rowsObjects.find(
          (row) => item['ID Network'] == row['ID NETWORK'],
        );
        if (find) {
          const dateString = date
            .toLocaleDateString('id-ID', {
              weekday: 'long',
              day: '2-digit',
              month: 'long',
              year: 'numeric',
              timeZone: 'Asia/Jakarta',
            })
            .toString();
          rows[find.index].assign({
            M: ` ${dateString}`,
          });
          await rows[find.index].save();
        }
      });
      return {};
    }

    if (subject.includes('pelaporan hasil audit')) {
      const OFFSET = 3;
      await doc.sheetsByTitle[stringMonth].loadHeaderRow(OFFSET);
      const table = this.outlookService.tableToJson({
        headerCount: 1,
        body,
      });
      const rowsObjects = rows.map((row) => {
        return { ...row.toObject(), index: row.rowNumber - OFFSET - 1 };
      });
      table.forEach(async (item) => {
        const find = rowsObjects.find(
          (row) => item['ID Network'] == row['ID NETWORK'],
        );
        if (find) {
          rows[find.index].assign({ O: item['ACTUAL CLOSING'] });
          await rows[find.index].save();
        }
      });
      return table;
    }

    return {
      message: 'Nothing.',
    };
  }

  @Post('create_header')
  async createHeader() {
    const doc = await this.spreadSheetService.doc();
    const sheet = await doc.addSheet({
      title:
        'Bulan ' +
        new Date().toLocaleDateString('id-ID', {
          month: 'long',
        }),
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
