import { Module } from '@nestjs/common';
import { OutlookService } from 'src/services/outlook.service';
import { HttpModule } from '@nestjs/axios';
import { SpreadsheetService } from 'src/services/spreadsheet.service';
import { NotificationController } from 'src/controllers/notification.controllers';

@Module({
  imports: [HttpModule],
  providers: [OutlookService, SpreadsheetService],
  controllers: [NotificationController],
})
export class NotificationModule {}
