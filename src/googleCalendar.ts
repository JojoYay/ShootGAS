import { ScriptProps } from './scriptProps';

export class GoogleCalendar {
    public syncAllToCalendar = () => {
        const setting: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(ScriptProps.instance.settingSheet);
        const calendarSheet: GoogleAppsScript.Spreadsheet.Sheet | null = setting.getSheetByName('calendar');

        if (!calendarSheet) {
            console.log('Calendar情報が見つからず、Google Calendarの同期ができませんでした。');
            throw new Error('Calendar情報が見つからず、Google Calendarの同期ができませんでした。');
        }

        const calendar = CalendarApp.getCalendarById(ScriptProps.instance.calendarId);
        // console.log(calendar);
        const dataRange = calendarSheet.getDataRange();
        const values = dataRange.getValues();

        for (let i = 1; i < values.length; i++) {
            // 1行目はヘッダーなのでスキップ
            const row = values[i];
            const title = row[2]; // event_name
            const startDate = new Date(row[3]); // start_datetime
            const endDate = new Date(row[4]); // end_datetime
            const eventId = row[11];
            const options = {
                description: row[6], // remark
                location: row[5],
            };

            this.syncEvent(eventId, calendar, title, startDate, endDate, options, i + 1, calendarSheet); // i + 1 は行番号
        }
    };

    // private updateEvent = (eventId:string, calendar:GoogleAppsScript.Calendar.Calendar, title:string, startDate:Date, endDate:Date, options:{description:string}):GoogleAppsScript.Calendar.CalendarEvent => {
    //     // カレンダーアイテム更新
    //     let event:GoogleAppsScript.Calendar.CalendarEvent = calendar.getEventById(eventId);
    //     if (event) {
    //         event.setTime(startDate, endDate);
    //         event.setDescription(options.description);
    //         event.setTitle(title);
    //     }
    //     return event;
    // }

    // private createEvent = (eventId:string, calendar:GoogleAppsScript.Calendar.Calendar, title:string, startDate:Date, endDate:Date, options:{description:string}):GoogleAppsScript.Calendar.CalendarEvent => {
    //     let event = calendar.createEvent(title, startDate, endDate, options);
    //     return event;
    // }

    public syncEvent = (
        eventId: string,
        calendar: GoogleAppsScript.Calendar.Calendar,
        title: string,
        startDate: Date,
        endDate: Date,
        options: { description: string; location: string },
        rowIndex: number,
        calendarSheet: GoogleAppsScript.Spreadsheet.Sheet
    ) => {
        if (eventId === '') {
            const eventIdCell = calendarSheet.getRange(rowIndex, 12);
            const event = calendar.createEvent(title, startDate, endDate, options);
            eventIdCell.setValue(event.getId());
        } else {
            // カレンダーアイテム更新
            const event: GoogleAppsScript.Calendar.CalendarEvent = calendar.getEventById(eventId);
            console.log(event);
            try {
                event.setTime(startDate, endDate);
                event.setDescription(options.description);
                event.setLocation(options.location);
                event.setTitle(title);
            } catch (error) {
                console.error('イベントの更新中にエラーが発生しました:', error);
            }
        }
    };
}
