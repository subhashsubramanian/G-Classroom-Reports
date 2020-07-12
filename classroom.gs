function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Google Classroom Reports');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Generates a sheet containing the summary of all courses with some basic stats.
 */
function populateCourseSummary() {
  var classroomSS = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var courses = getAllCourses();
    
    var summarySheet = createSheet(classroomSS, "Course Summary"); 
    summarySheet.appendRow([
      "COURSE ID", "NAME", "NUMBER OF STUDENTS", 
      "COURSEWORK COMPLETION RATE", "AVERAGE COMPLETION TIME (days)", 
      "AVERAGE GRADING TIME (days)", "AVERAGE GRADE", "MAXIMUM GRADE", 
      "MINIMUM GRADE", "MEDIAN GRADE"]);
    formatHeader(summarySheet, "A1:J1");
    
    
    if (courses && courses.length > 0) {
      for (var i = 0; i < courses.length; i++) {
        var course = courses[i];
        
        var response2 = Classroom.Courses.Students.list(course.id);
        var students = response2.students;
        
        var rate = getSubmissionRate(course.id, "-");
        var averageTime = getAverageSubmissionTime(course.id, "-");
        var averageGradeTime = getAverageGradeTime(course.id, "-");
        var averageGrade = getAverageGrade(course.id, "-");
        var maxGrade = getMaximumGrade(course.id, "-");
        var minGrade = getMinimumGrade(course.id, "-");
        var medianGrade = getMedianGrade(course.id, "-");
        
        summarySheet.appendRow([course.id, course.name, students.length, rate, 
                                averageTime, averageGradeTime, averageGrade, 
                                maxGrade, minGrade, medianGrade]);
      }
      formatTableRows(summarySheet, "A2:J" + (courses.length + 1));
    } else {
      return "No courses found."
    }
  }
  catch (err) {
    Logger.log(err);
    return "An error has occured while trying to get courses.";
  }
  return "Success."
}

/**
 * Generates a sheet for each of the courses and adds details about the course.
 */
function populateCourseDetails() {
  var classroomSS = SpreadsheetApp.getActiveSpreadsheet();
  try{
    var courses = getAllCourses();
    
    if (courses && courses.length > 0) {
      for (i = 0; i < courses.length; i++) {
        var course = courses[i];
        var courseSheet = createSheet(classroomSS, course.name);
        var courseWork = getAllCourseWork(course.id);
        populateCourseWorkList(course.id, courseWork, courseSheet);
        populateStudentGradeList(course.id, courseWork, courseSheet);
      }
    } else {
      return 'No courses found.';
    }
  }
  catch (err) {
    Logger.log(err);
    return "An error has occured while trying to get courses.";
  }
  return "Success."
}

/**
 * Lists all course work for a  given course.
 */
function populateCourseWorkList(courseId, courseWork, courseSheet) {
  courseSheet.appendRow([
      "TITLE", "COMPLETION RATE", "AVERAGE COMPLETION TIME (days)", 
      "AVERAGE GRADING TIME (days)", "AVERAGE GRADE", "MAXIMUM GRADE", 
      "MINIMUM GRADE", "MEDIAN GRADE"]);
  var headerIndex = courseSheet.getLastRow();
  formatHeader(courseSheet, "A" + headerIndex + ":H" + headerIndex);
  if (courseWork && courseWork.length > 0) {
    for (var j = 0; j < courseWork.length; j++) {
      Logger.log(courseWork[j].title);
      if (courseWork[j]) {
        var rate = getSubmissionRate(courseId, courseWork[j].id);
        var averageTime = getAverageSubmissionTime(courseId, courseWork[j].id);
        var averageGradeTime = getAverageGradeTime(courseId, courseWork[j].id);
        var averageGrade = getAverageGrade(courseId, courseWork[j].id);
        var maxGrade = getMaximumGrade(courseId, courseWork[j].id);
        var minGrade = getMinimumGrade(courseId, courseWork[j].id);
        var medianGrade = getMedianGrade(courseId, courseWork[j].id);
        courseSheet.appendRow([
          courseWork[j].title, rate, averageTime, averageGradeTime, 
          averageGrade, maxGrade, minGrade, medianGrade]);
      }
    }
    formatTableRows(courseSheet, "A" + (headerIndex + 1) + ":H" + (courseWork.length + 1));
  }
  else {
    Logger.log("No coursework found for course %s", courseId);
  }
}

/**
 * Lists all the students with their average grade.
 */
function populateStudentGradeList(courseId, courseWork, courseSheet) {
  courseSheet.appendRow([" "]);
  //courseSheet.appendRow(["STUDENT NAME","AVERAGE GRADE"]);
  
  // make array of all coursework titles
  var courseWorkAlltitles = [];
  for (var j = 0; j < courseWork.length; j++) {
    courseWorkAlltitles[j] = courseWork[j].title;
  }
  
    
  var headings = ["Student Name"];
  headings = headings.concat(["Total"]);
  headings = headings.concat(courseWorkAlltitles);
  
  //Logger.log(headings);
  
  //courseSheet.appendRow(["STUDENT NAME","AVERAGE GRADE"]);
  courseSheet.appendRow(headings);
  
  
  var headerIndex = courseSheet.getLastRow();
  var lastCol = "";
  // number of courseworks + student name + total = courseWork.length +2 
  lastCol = columnToLetter(courseWork.length +2);
  
  formatHeader(courseSheet, "A" + headerIndex + ":" + lastCol + headerIndex);
  
  var response2 = Classroom.Courses.Students.list(courseId);
  var students = response2.students;
  var studentNameAndGrades = [];
     
  if (students && students.length > 0) {
    for (z = 0; z < students.length; z++) {
      if (!(typeof students[z].userId === 'undefined')) {
        studentNameAndGrades[0] = students[z].profile.name.fullName
        //Browser.msgBox(getcourseWorkStudentGrade(courseId, courseWork, students[z].userId))
        studentNameAndGrades = studentNameAndGrades.concat(getcourseWorkStudentGrade(courseId, courseWork, students[z].userId))
        courseSheet.appendRow(studentNameAndGrades);
        studentNameAndGrades = [];
        //getAverageStudentGrade(courseId, courseWork, students[z].userId)]);
      }
    }
  }
  //formatTableRows(courseSheet, "A" + (headerIndex + 1) + ":H" + (headerIndex + students.length));
  
   formatTableRows(courseSheet, "A" + (headerIndex + 1) + ":" + lastCol + (headerIndex + students.length));
   
  if (courseWork && courseWork.length > 0) {
    makegradesHistogram(headerIndex + 1, headerIndex + students.length, courseSheet)  
    
    //populateGradesHistogram(courseSheet, headerIndex + 1, headerIndex + students.length, students.length);
    // populateStats(courseSheet, headerIndex, headerIndex + students.length);
  }
}



function makegradesHistogram(startrow, endrow, courseSheet) {
  //var spreadsheet = SpreadsheetApp.getActive();
  courseSheet.getRange('B' + startrow + ':B' + endrow).activate();
  //var sheet = spreadsheet.getActiveSheet();
  chart = courseSheet.newChart()
  .asHistogramChart()
  .addRange(courseSheet.getRange('B' + startrow + ':B' + endrow))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('title', 'Histogram of Marks')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setYAxisTitle('Total')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setPosition(endrow+3, 1, 0, 0)
  .build();
  
  courseSheet.insertChart(chart);
  //var charts = sheet.getCharts();
  //chart = charts[charts.length - 1];
  //sheet.removeChart(chart);
  
};



/**
 * Populates the grade distribution histogram.
 */
function populateGradesHistogram(courseSheet, startIndex, endIndex, length) {
    
  // var lastCol = courseSheet.getLastColumn();
  
  var dataRange = courseSheet.getRange("B"+ (startIndex+1) + ":B" + (endIndex+1));
  
  var chart = courseSheet.newChart()
  .setChartType(Charts.ChartType.HISTOGRAM)
  .addRange(dataRange)
  .setPosition(startIndex + 3, 4, 0, 0)
  .setOption('title', 'Grade Histogram')
  .asHistogramChart()
  .setColors(["#729460"])
  .setOption('chartArea', '{width: 100, height: 100}')
  .setTitleTextStyle(Charts.newTextStyle().setColor('#212b33').setFontSize(15).build())
  .setBackgroundColor("#d6d6d6")
  .build();
  
  courseSheet.insertChart(chart);
}

/**
 * Populates stats (average, median, mode, stdev) from the grades.
 */ 
function populateStats(courseSheet, startIndex, endIndex) {
  var values = [["MEAN","MEDIAN","MODE","STDEVP"]];
  
  var headerRange = "E"+ startIndex + ":H"+ startIndex;
  courseSheet.getRange(headerRange).setValues(values);
  
  formatHeader(courseSheet, headerRange);
  
  var rangeString = "B"+ (startIndex+1) + ":B" + (endIndex+1);
  
  var formulae = [[
    "=AVERAGE(" + rangeString + ")",
    "=MEDIAN("+ rangeString + ")",
    "=MODE("+ rangeString + ")",
    "=STDEVP("+ rangeString + ")"
  ]];
  
  cellsRange = "E"+ (startIndex + 1) + ":H"+ (startIndex + 1);
  courseSheet.getRange(cellsRange).setFormulas(formulae);
  
  formatTableRows(courseSheet, cellsRange);
}
