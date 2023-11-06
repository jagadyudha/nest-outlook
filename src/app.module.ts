import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
// import { NotificationModule } from './modules/notification.module';
import { SpreadsheetModule } from './modules/spreadsheet.module';
@Module({
  imports: [ConfigModule.forRoot(), SpreadsheetModule],
})
export class AppModule {}
