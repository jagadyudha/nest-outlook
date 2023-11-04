import { Module } from '@nestjs/common';
import { OutlookService } from 'src/services/outlook.service';
import { HttpModule } from '@nestjs/axios';
import { SpreadsheetService } from 'src/services/spreadsheet.service';
import { SpreadsheetController } from 'src/controllers/spreadsheet.controller';

@Module({
  imports: [HttpModule],
  providers: [OutlookService, SpreadsheetService],
  controllers: [SpreadsheetController],
})
export class NotificationModule {}
