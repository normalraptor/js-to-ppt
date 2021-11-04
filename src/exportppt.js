/* eslint-disable no-unused-vars */
import pptxgen from "pptxgenjs";
function ExportPPT(){
    let pptx =  new pptxgen();
    let slideOverview  = pptx.addSlide();
    let slide2 = pptx.addSlide();
    let slide3 = pptx.addSlide();
    let slide4 = pptx.addSlide();
    let slide5 = pptx.addSlide();
    let slide6 = pptx.addSlide();
    let slide7 = pptx.addSlide();
    let slideChart = []
    
    for (let i = 0; i<40; i++){
      slideChart[i] = pptx.addSlide();
    }

    let perfomanceReportComment = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
    let resourceReportComment = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
    let benchmarkReportComment = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
/*
~~~~OVERVIEW~~~~~
*/
  slideOverview.addText(
    "Report Overview",
    {
      x: 1,
      y: 3,
      w: 6.3,
      h: 1,
      fontSize: 60,
      align: "center"
    });

/*
~~~~SLIDE 2 PERFORMANCE REPORT~~~~~
*/
slide2.addText(
  "Performance Report",
  {
    x: 0.2,
    y: 0.25,
    w: 9.45,
    h: 0.64,
    fontSize: 28,
    bold: true,
    align: "center",
  });


//Row 1 Col 1    
slide2.addText(
  "Judul Gambar 1",
  {
    x: 0.5,
    y: 1,
    w: 3.075,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 8}
  });
        
slide2.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 1.17,
    w: 3.105,
    h: 1.885,
    line: { color: "000000", width: 1 },
  });
          
slide2.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 1.20,
    w: 3.075,
    h: 1.845,
  });


//Row 2 col 1
slide2.addText(
  "Judul Gambar 3",
  {
    x: 0.5,
    y: 3.145,
    w: 3.075,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 10}
  });
                   
slide2.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 3.315,
    w: 3.115,
    h: 1.885,
    line: { color: "000000", width: 1 },
  });
          
slide2.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 0.5,
    y: 3.345,
    w: 3.075,
    h: 1.845,
  });

//Row 1 col 2
slide2.addText(
  "Judul Gambar 2",
  {
    x: 3.725,
    y: 1,
    w: 3.075,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 9}
  });
           
slide2.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 3.725,
    y: 1.17,
    w: 3.115,
    h: 1.885,
    line: { color: "000000", width: 1 },
  });
        
slide2.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 3.725,
    y: 1.20,
    w: 3.075,
    h: 1.845,
  });

//Row 2 col 2
slide2.addText(
  "Judul Gambar 4",
  {
    x: 3.725,
    y: 3.145,
    w: 3.075,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 11}
  });
           
slide2.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 3.725,
    y: 3.315,
    w: 3.115,
    h: 1.885,
    line: { color: "000000", width: 1 },
  });
          
slide2.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 3.725,
    y: 3.345,
    w: 3.075,
    h: 1.845,
  });
  

//Comment column
slide2.addText(
  "PERFORMANCE",
  {
    x: 6.95,
    y: 1,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    color: "ffffff",
    fill: {color: "24207A" },
  });

slide2.addText(
  "Comments",
  {
    x: 6.95,
    y: 1.3,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    fill: {color: "F1F1F1" },
  });

slide2.addTable(
  [
    [
      {
        text: perfomanceReportComment
      }
    ]
  ],
  {
    x: 6.95,
    y: 1.6,
    w: 3,
    h: 3,
    fontSize: 12,
    border: {color: "9c9c9c"}
  });
    

/*
~~~~SLIDE 3 RESOURCE REPORT PART 1 OF 3~~~~~
*/
  slide3.addText(
    "Resources Report",
    {
      x: 0.2,
      y: 0.25,
      w: 9.45,
      h: 0.64,
      fontSize: 28,
      bold: true,
      align: "center",
    });


//Row 1 Col 1    
  slide3.addText(
    "Judul Gambar 1",
    {
      x: 0.5,
      y: 1,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 12}
    });
          
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 0.5,
      y: 1.17,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
      x: 0.5,
      y: 1.20,
      w: 2,
      h: 1.2,
    });


//Row 2 col 1
  slide3.addText(
    "Judul Gambar 4",
    {
      x: 0.5,
      y: 2.47,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 15}
    });
                     
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 0.5,
      y: 2.64,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
      x: 0.5,
      y: 2.67,
      w: 2,
      h: 1.2,
    });

//row 3 col 1
  slide3.addText(
    "Judul Gambar 7",
    {
      x: 0.5,
      y: 4.04,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 18}
    });
            
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 0.5,
      y: 4.21,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
      x: 0.5,
      y: 4.24,
      w: 2,
      h: 1.2,
    });

//Row 1 col 2
  slide3.addText(
    "Judul Gambar 2",
    {
      x: 2.65,
      y: 1,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 13}
    });
             
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 2.65,
      y: 1.17,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
          
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
      x: 2.68,
      y: 1.20,
      w: 2,
      h: 1.2,
    });

//Row 2 col 2
  slide3.addText(
    "Judul Gambar 5",
    {
      x: 2.65,
      y: 2.47,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 16}
    });
             
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 2.65,
      y: 2.64,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
      x: 2.68,
      y: 2.67,
      w: 2,
      h: 1.2,
    });
    
//Row 3 col 2
  slide3.addText(
    "Judul Gambar 8",
    {
      x: 2.65,
      y: 4.04,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 19}
    });
            
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 2.65,
      y: 4.21,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
      x: 2.68,
      y: 4.24,
      w: 2,
      h: 1.2,
    });
          
//Row 1 col 3
  slide3.addText(
    "Judul Gambar 3",
    {
      x: 4.8,
      y: 1,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 14}
    });
              
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 4.8,
      y: 1.17,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
      x: 4.83,
      y: 1.20,
      w: 2,
      h: 1.2,
    });
    
//Row 2 col 3
  slide3.addText(
    "Judul Gambar 6",
    {
      x: 4.8,
      y: 2.47,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 17}
    });
             
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 4.8,
      y: 2.64,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
      x: 4.83,
      y: 2.67,
      w: 2,
      h: 1.2,
    });

//Row 3 col 3
  slide3.addText(
    "Judul Gambar 9",
    {
      x: 4.8,
      y: 4.04,
      w: 2,
      h: 0.1,
      fontSize: 14,
      align: "center",
      hyperlink: {slide: 20}
    });
            
  slide3.addShape(
    pptx.shapes.RECTANGLE,
    {
      x: 4.8,
      y: 4.21,
      w: 2.04,
      h: 1.24,
      line: { color: "000000", width: 1 },
    });
            
  slide3.addImage(
    {
      path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",    
      x: 4.83,
      y: 4.24,
      w: 2,
      h: 1.2,
    });
  
//Comment column
  slide3.addText(
    "RESOURCES",
    {
      x: 6.95,
      y: 1,
      w: 3,
      h: 0.3,
      fontSize: 16,
      align: "center",
      bold: true,
      color: "ffffff",
      fill: {color: "24207A" },
    });
  
  slide3.addText(
    "Comments",
    {
      x: 6.95,
      y: 1.3,
      w: 3,
      h: 0.3,
      fontSize: 16,
      align: "center",
      bold: true,
      fill: {color: "F1F1F1" },
    });
  
  slide3.addTable(
    [
      [
        {
          text: resourceReportComment
        }
      ]
    ],
    {
      x: 6.95,
      y: 1.6,
      w: 3,
      h: 3,
      fontSize: 12,
      border: {color: "9c9c9c"}
    });
  


/*
~~~~SLIDE 4 RESOURCE REPORT PART 2 OF 3~~~~~
*/
slide4.addText(
  "Resources Report Cont.",
  {
    x: 0.2,
    y: 0.25,
    w: 9.45,
    h: 0.64,
    fontSize: 28,
    bold: true,
    align: "center",
  });


//Row 1 Col 1    
slide4.addText(
  "Judul Gambar 10",
  {
    x: 0.5,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 21}
  });
        
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 1.20,
    w: 2,
    h: 1.2,
  });


//Row 2 col 1
slide4.addText(
  "Judul Gambar 13",
  {
    x: 0.5,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 24}
  });
                   
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 0.5,
    y: 2.67,
    w: 2,
    h: 1.2,
  });

//row 3 col 1
slide4.addText(
  "Judul Gambar 16",
  {
    x: 0.5,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 27}
  });
          
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 4.24,
    w: 2,
    h: 1.2,
  });

//Row 1 col 2
slide4.addText(
  "Judul Gambar 11",
  {
    x: 2.65,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 22}
  });
           
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
        
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 2.68,
    y: 1.20,
    w: 2,
    h: 1.2,
  });

//Row 2 col 2
slide4.addText(
  "Judul Gambar 14",
  {
    x: 2.65,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 25}
  });
           
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 2.68,
    y: 2.67,
    w: 2,
    h: 1.2,
  });
  
//Row 3 col 2
slide4.addText(
  "Judul Gambar 17",
  {
    x: 2.65,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 28}
  });
          
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 2.68,
    y: 4.24,
    w: 2,
    h: 1.2,
  });
        
//Row 1 col 3
slide4.addText(
  "Judul Gambar 12",
  {
    x: 4.8,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 23}
  });
            
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 4.83,
    y: 1.20,
    w: 2,
    h: 1.2,
  });
  
//Row 2 col 3
slide4.addText(
  "Judul Gambar 15",
  {
    x: 4.8,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 26}
  });
           
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 4.83,
    y: 2.67,
    w: 2,
    h: 1.2,
  });

//Row 3 col 3
slide4.addText(
  "Judul Gambar 18",
  {
    x: 4.8,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 29}
  });
          
slide4.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide4.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",    
    x: 4.83,
    y: 4.24,
    w: 2,
    h: 1.2,
  });

//Comment column
slide4.addText(
  "RESOURCES",
  {
    x: 6.95,
    y: 1,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    color: "ffffff",
    fill: {color: "24207A" },
  });

slide4.addText(
  "Comments",
  {
    x: 6.95,
    y: 1.3,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    fill: {color: "F1F1F1" },
  });

slide4.addTable(
  [
    [
      {
        text: resourceReportComment
      }
    ]
  ],
  {
    x: 6.95,
    y: 1.6,
    w: 3,
    h: 3,
    fontSize: 12,
    border: {color: "9c9c9c"}
  });


/*
~~~~SLIDE 5 RESOURCE REPORT PART 3 OF 3~~~~~
*/
        
slide5.addText(
  "Resources Report Cont.",
  {
    x: 0.2,
    y: 0.25,
    w: 9.45,
    h: 0.64,
    fontSize: 28,
    bold: true,
    align: "center",
  });


//Row 1 Col 1    
slide5.addText(
  "Judul Gambar 19",
  {
    x: 0.5,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 30}
  });
        
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 1.20,
    w: 2,
    h: 1.2,
  });


//Row 2 col 1
slide5.addText(
  "Judul Gambar 22",
  {
    x: 0.5,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 33}
  });
                   
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 0.5,
    y: 2.67,
    w: 2,
    h: 1.2,
  });

//row 3 col 1
slide5.addText(
  "Judul Gambar 25",
  {
    x: 0.5,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 36}
  });
          
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 4.24,
    w: 2,
    h: 1.2,
  });

//Row 1 col 2
slide5.addText(
  "Judul Gambar 20",
  {
    x: 2.65,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 31}
  });
           
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
        
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 2.68,
    y: 1.20,
    w: 2,
    h: 1.2,
  });

//Row 2 col 2
slide5.addText(
  "Judul Gambar 23",
  {
    x: 2.65,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 34}
  });
           
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 2.68,
    y: 2.67,
    w: 2,
    h: 1.2,
  });
  
//Row 3 col 2
slide5.addText(
  "Judul Gambar 26",
  {
    x: 2.65,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 37}
  });
          
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 2.68,
    y: 4.24,
    w: 2,
    h: 1.2,
  });
        
//Row 1 col 3
slide5.addText(
  "Judul Gambar 21",
  {
    x: 4.8,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 32}
  });
            
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 4.83,
    y: 1.20,
    w: 2,
    h: 1.2,
  });
  
//Row 2 col 3
slide5.addText(
  "Judul Gambar 24",
  {
    x: 4.8,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 35}
  });
           
slide5.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide5.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 4.83,
    y: 2.67,
    w: 2,
    h: 1.2,
  });


//Comment column
slide5.addText(
  "RESOURCES",
  {
    x: 6.95,
    y: 1,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    color: "ffffff",
    fill: {color: "24207A" },
  });

slide5.addText(
  "Comments",
  {
    x: 6.95,
    y: 1.3,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    fill: {color: "F1F1F1" },
  });

slide5.addTable(
  [
    [
      {
        text: resourceReportComment
      }
    ]
  ],
  {
    x: 6.95,
    y: 1.6,
    w: 3,
    h: 3,
    fontSize: 12,
    border: {color: "9c9c9c"}
  });


/*
~~~~SLIDE 6 BENCHMARK REPORT~~~~~
*/
slide6.addText(
  "Benchmark Report",
  {
    x: 0.2,
    y: 0.25,
    w: 9.45,
    h: 0.64,
    fontSize: 28,
    bold: true,
    align: "center",
  });


//Row 1 Col 1    
slide6.addText(
  "Judul Gambar 1",
  {
    x: 0.5,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 38}
  });
        
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 1.20,
    w: 2,
    h: 1.2,
  });


//Row 2 col 1
slide6.addText(
  "Judul Gambar 4",
  {
    x: 0.5,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 41}
  });
                   
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 0.5,
    y: 2.67,
    w: 2,
    h: 1.2,
  });

//row 3 col 1
slide6.addText(
  "Judul Gambar 7",
  {
    x: 0.5,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 44}
  });
          
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 0.5,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
    x: 0.5,
    y: 4.24,
    w: 2,
    h: 1.2,
  });

//Row 1 col 2
slide6.addText(
  "Judul Gambar 2",
  {
    x: 2.65,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 39}
  });
           
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
        
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU", 
    x: 2.68,
    y: 1.20,
    w: 2,
    h: 1.2,
  });

//Row 2 col 2
slide6.addText(
  "Judul Gambar 5",
  {
    x: 2.65,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 42}
  });
           
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 2.68,
    y: 2.67,
    w: 2,
    h: 1.2,
  });
  
//Row 3 col 2
slide6.addText(
  "Judul Gambar 8",
  {
    x: 2.65,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 45}
  });
          
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 2.65,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 2.68,
    y: 4.24,
    w: 2,
    h: 1.2,
  });
        
//Row 1 col 3
slide6.addText(
  "Judul Gambar 3",
  {
    x: 4.8,
    y: 1,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 40}
  });
            
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 1.17,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 4.83,
    y: 1.20,
    w: 2,
    h: 1.2,
  });
  
//Row 2 col 3
slide6.addText(
  "Judul Gambar 6",
  {
    x: 4.8,
    y: 2.47,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 43}
  });
           
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 2.64,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",
    x: 4.83,
    y: 2.67,
    w: 2,
    h: 1.2,
  });

//Row 3 col 3
slide6.addText(
  "Judul Gambar 9",
  {
    x: 4.8,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 46}
  });
          
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 4.8,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",    
    x: 4.83,
    y: 4.24,
    w: 2,
    h: 1.2,
  });

//extra below comment
slide6.addText(
  "Judul Gambar 10",
  {
    x: 7.45,
    y: 4.04,
    w: 2,
    h: 0.1,
    fontSize: 14,
    align: "center",
    hyperlink: {slide: 47}
  });
          
slide6.addShape(
  pptx.shapes.RECTANGLE,
  {
    x: 7.45,
    y: 4.21,
    w: 2.04,
    h: 1.24,
    line: { color: "000000", width: 1 },
  });
          
slide6.addImage(
  {
    path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",    
    x: 7.48,
    y: 4.24,
    w: 2,
    h: 1.2,
  });

//Comment column
slide6.addText(
  "BENCHMARK",
  {
    x: 6.95,
    y: 1,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    color: "ffffff",
    fill: {color: "24207A" },
  });

slide6.addText(
  "Comments",
  {
    x: 6.95,
    y: 1.3,
    w: 3,
    h: 0.3,
    fontSize: 16,
    align: "center",
    bold: true,
    fill: {color: "F1F1F1" },
  });

slide6.addTable(
  [
    [
      {
        text: benchmarkReportComment
      }
    ]
  ],
  {
    x: 6.95,
    y: 1.6,
    w: 3,
    h: 2.28,
    fontSize: 11,
    border: {color: "9c9c9c"}
  });


/*
~~~~SLIDE 7 CHART DETAILS~~~~~
*/
  slide7.addText(
    "Chart Details",
    {
      x: 1,
      y: 3,
      w: 6.3,
      h: 1,
      fontSize: 60,
      align: "center"
    });


/*
~~~~SLIDE 8 - 47 FULL CHARTS~~~~~
*/
for (let i = 0; i<40; i++){
  if(i<4)
  {
    slideChart[i].addText(
      "Performance Chart",
      {
        x: 0.2,
        y: 0.25,
        w: 9.45,
        h: 0.64,
        fontSize: 28,
        bold: true,
        align: "center",
      }
    );

    slideChart[i].addText(
      "Kembali",
      {
        x: 0,
        y: 0.1,
        w: 1,
        h: 0.1,
        fontSize: 14,
        color: "2339DB",
        bold: true,
        align: "left",
        hyperlink: {slide: 2},
      }
    );

    slideChart[i].addShape(
      pptx.shapes.RECTANGLE,
      {
        x: 1.5,
        y: 1.17,
        w: 7.04,
        h: 4.24,
        line: { color: "000000", width: 1 },
      });
              
    slideChart[i].addImage(
      {
        path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
        x: 1.5,
        y: 1.20,
        w: 7,
        h: 4.2,
      });
    
  }else
  {
    if(i<30)
    {
      slideChart[i].addText(
        "Resource Chart",
        {
          x: 0.2,
          y: 0.25,
          w: 9.45,
          h: 0.64,
          fontSize: 28,
          bold: true,
          align: "center",
        }
      );
  
      slideChart[i].addText(
        "Kembali",
        {
          x: 0,
          y: 0.1,
          w: 1,
          h: 0.1,
          fontSize: 14,
          color: "2339DB",
          bold: true,
          align: "left",
          hyperlink: {slide: i<13 ? 3 : i<22 ? 4 : 5},
        }
      );
  
      slideChart[i].addShape(
        pptx.shapes.RECTANGLE,
        {
          x: 1.5,
          y: 1.17,
          w: 7.04,
          h: 4.24,
          line: { color: "000000", width: 1 },
        });
                
      slideChart[i].addImage(
        {
          path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
          x: 1.5,
          y: 1.20,
          w: 7,
          h: 4.2,
        });

    }else
    {
      slideChart[i].addText(
        "Benchmark Chart",
        {
          x: 0.2,
          y: 0.25,
          w: 9.45,
          h: 0.64,
          fontSize: 28,
          bold: true,
          align: "center",
        }
      );
  
      slideChart[i].addText(
        "Kembali",
        {
          x: 0,
          y: 0.1,
          w: 1,
          h: 0.1,
          fontSize: 14,
          color: "2339DB",
          bold: true,
          align: "left",
          hyperlink: {slide: 6},
        }
      );
  
      slideChart[i].addShape(
        pptx.shapes.RECTANGLE,
        {
          x: 1.5,
          y: 1.17,
          w: 7.04,
          h: 4.24,
          line: { color: "000000", width: 1 },
        });
                
      slideChart[i].addImage(
        {
          path: "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRmYlXbtDtjgFAic1EYkxRE05lvAuuLbnuNQA&usqp=CAU",  
          x: 1.5,
          y: 1.20,
          w: 7,
          h: 4.2,
        });
    }
  }
}


    pptx.writeFile({ fileName: 'NSEM-Report-Analytic-.pptx' });
}
  
export default ExportPPT;