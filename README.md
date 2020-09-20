# small-project2
## Tesseract(Open Source)를 활용한 검수 자동화 프로그램 개발    
+ Input Data : capturefile_list.xlsx    
  + capturefile_list.xlsx : 실제 디스크 속 파일의 속성과, 복사된 파일의 속성이 하나의 이미지로 저장되어있는 파일들의 위치가 기록되어있다.    
+ Output Data : Result.xlsx    
  + Result.xlsx : 아래의 조건을 통과하는 sample의 Predict column에 'O' 마킹이 되어있다.    
  
+ 조건    
1. 이미 검수가 완료되어있지 않은가    
2. 해당 위치에 파일이 있는가    
3. 이미지의 형식이 폴더의 속성의 형식인가(정해진 규격에 일치하는가)    
4. 읽어온 이미지의 텍스트를 추출했을 때, '파일'(또는 '폴더')에 매칭되는 숫자가 두 개이고, 서로 값이 같으며 엑셀의 적힌 count와의 숫자와도 같은가     
