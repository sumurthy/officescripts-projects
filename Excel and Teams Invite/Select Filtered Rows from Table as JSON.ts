function main(workbook: ExcelScript.Workbook): InterviewInvite[] {

  const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue() as number;
  const sheet = workbook.getWorksheet('Interviews');
  const table = sheet.getTables()[0];
  table.getColumnByName('Start time1').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("[$-en-US]m/d/yyyy h:mm AM/PM;@");
  table.getColumnByName('Start time2').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("[$-en-US]m/d/yyyy h:mm AM/PM;@");
  const dataRows: (string | number | boolean)[][] = table.getRange().getTexts();
  // or
  // let dataRows = sheet.getUsedRange().getValues();
  const selectedRows = dataRows.filter((row, i) => {
    // Select header row and any data row with the status column equal to approach value
    return (row[1] === 'FALSE' || i === 0)
  })
  const recordDetails: RecordDetail[] = returnObjectFromValues(selectedRows as string[][]);
  const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
  console.log(JSON.stringify(inviteRecords));
  return inviteRecords;
}

/**
 * This helper funciton converts table values into an object array.
 */
function returnObjectFromValues(values: string[][]): RecordDetail[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as RecordDetail[];
}

/**
 * Generate interview records by selecting required columns
 * @param records Input records
 * @param mins Number of minutes to add to the start date-time
 */
function generateInterviewRecords(records: RecordDetail[], mins: number): InterviewInvite[] {
  const interviewinvites: InterviewInvite[] = []

  records.forEach((record) => {
    // Interviewer 1    
    // If the start date-time is greather than current date-time, add to output records
    if ((new Date(record['Start time1'])) > new Date()) {
      let startTime = new Date(record['Start time2']).toISOString();
      // compute the finish time of the meeting
      let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
      interviewinvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer1,
        InterviewerEmail: record['Interviewer1 email'],
        StartTime: startTime, 
        FinishTime: finishTime 
      })
    }
    // Interviewer 2 
    // If the start date-time is greather than current date-time, add to output records
    if ((new Date(record['Start time2'])) > new Date()) {
      let startTime = new Date(record['Start time2']).toISOString();
      // compute the finish time of the meeting
      let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
      interviewinvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer2,
        InterviewerEmail: record['Interviewer2 email'],
        StartTime: startTime,
        FinishTime: finishTime 
      })
    }
  })
  return interviewinvites;
}

/**
 * Add minutes to start date-time
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date time
 */
function addMins(startDateTime: Date, mins: number) {
  return new Date(startDateTime.getTime() + mins * 60 * 1000);
}

// Basic key-value pair object
interface BasicObj {
  [key: string]: string | number | boolean
}

// Input record that matches the table data
interface RecordDetail extends BasicObj {
  ID: string
  'Invite to interview': string
  Candidate: string
  'Candidate email': string
  'Candidate contact': string
  Interviewer1: string
  'Interviewer1 email': string
  Interviewer2: string
  'Interviewer2 email': string
  'Start time1': string
  'Start time2': string
}

// Output record
interface InterviewInvite extends BasicObj {
  ID: string
  Candidate: string
  CandidateEmail: string
  CandidateContact: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
}