Build the project:
mvn clean package

Run the application:
mvn exec:java -Dexec.mainClass=excel.ExcelWriter -Dexec.arguments="__OUTPUT_FOLDER__"
