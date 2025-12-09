msg-runner project
==================

What you got:
- A Spring Boot web app that serves a simple HTML UI at / (http://localhost:8080)
- A controller that triggers MA_MSG_Suite_INB.MSG_MAIN (a runnable stub by default)
- Your original MSG_MAIN file included as MSG_MAIN_original.java inside the MA_MSG_Suite_INB package
  (replace the stub MSG_MAIN.java with your original if you want to run your real flow).
- Dockerfile and Maven pom.xml included.

How to run locally (no Docker):
1. Open terminal, go to project root (where pom.xml is).
2. mvn clean package
3. java -jar target/msg-runner-1.0.0.jar
4. Open http://localhost:8080

How to run with Docker:
1. docker build -t msg-runner .
2. docker run -p 8080:8080 msg-runner
3. Open http://localhost:8080

Notes:
- The project includes a safe, compiling stub for MA_MSG_Suite_INB.MSG_MAIN (so the project builds out-of-the-box).
- Your original MSG_MAIN (the file you shared) is present as MSG_MAIN_original.java in the same package.
  If you want to run your real code, replace src/main/java/MA_MSG_Suite_INB/MSG_MAIN.java with your original
  MSG_MAIN content, and ensure any other referenced classes are present in the project classpath.
- If your real MSG_MAIN uses Selenium & ChromeDriver, when running on a headless server or Docker you'll need
  to configure ChromeOptions to run in headless mode and ensure Chrome/Chromium and required libs are available.
