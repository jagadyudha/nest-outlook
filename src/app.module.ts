import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
// import { NotificationModule } from './modules/notification.module';
import { SpreadsheetController } from './controllers/spreadsheet.controller';
@Module({
  imports: [ConfigModule.forRoot(), SpreadsheetController],
})
export class AppModule {}
