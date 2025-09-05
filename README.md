## Overview
MFRS is a Spring Boot based web application for generating monthly financial reports (Jasper reports + charts). It uses Thymeleaf for server-side pages and Spring Security for authentication/authorization. The project is packaged as a WAR and is designed to be deployed to a servlet container (e.g., Apache Tomcat). It also contains static client assets and generated report templates.

## Quick facts / What you’ll find in this repo
- Project root: `d:\MFRS`
- Main entry-point class: `com.rcs.mfrs.MfrsApplication`
- Primary configuration: `src/main/resources/application.properties`
- Templates: `src/main/resources/templates`
- Static assets: `src/main/resources/static` (css, js, images)
- Reports: `src/main/resources/company.jrxml` (JasperReports templates)
- Build artifact (generated): `target/mfrs-0.0.1-SNAPSHOT.war`

## Technical specifications
- Java: 17 (see `<properties><java.version>17</java.version>` in `pom.xml`)
- Spring Boot parent: 2.7.5
- Packaging: WAR (`<packaging>war</packaging>`)
- Build tool: Maven (mvn wrapper included: `mvnw`, `mvnw.cmd`)
- Lombok: 1.18.28 (annotation processing configured)
- Notable dependencies (from `pom.xml`):
  - spring-boot-starter-web
  - spring-boot-starter-data-jpa
  - spring-boot-starter-thymeleaf
  - spring-boot-starter-security
  - spring-boot-starter-validation
  - mysql-connector-j (8.0.33)
  - jasperreports (6.21.3)
  - jfreechart (1.5.4)
  - thymeleaf-extras-springsecurity5
  - spring-boot-devtools (runtime, optional)
- Maven Compiler Plugin: 3.11.0 (source/target set to 17)

> Note: The `spring-boot-starter-tomcat` dependency is declared with `<scope>provided</scope>` which implies the WAR is intended to be deployed to an external servlet container (such as Tomcat) rather than using embedded Tomcat at runtime.

## Architecture & package layout
The project follows a typical Spring Boot MVC structure:

- `src/main/java/com/rcs/mfrs/` - application classes
  - `MfrsApplication.java` - main Spring Boot bootstrap class
  - Additional controllers, services, repositories, DTOs, entities and security classes are expected under this package (see `src/main/java/com/rcs` in repo)

- `src/main/resources/`
  - `application.properties` - application configuration (DB, JPA properties)
  - `company.jrxml` - JasperReports template
  - `templates/` - Thymeleaf HTML views (`index.html`, `home.html`, `login.html`, etc.)
  - `static/` - static assets (CSS, JS, images). Chart libraries and custom JS are included here.

- `target/` - Maven build output including the WAR and expanded classes/resources used by deployment.

High-level responsibilities
- Controllers: handle web routes and return Thymeleaf views or JSON.
- Services: business logic, data transformation, and report orchestration.
- Repositories (Spring Data JPA): persistence layer connecting to MySQL.
- Templates & Reports: UI pages + Jasper report generation (PDF/Excel) using `company.jrxml` and compiled jasper files.

## Main application behaviour
- The main class (`com.rcs.mfrs.MfrsApplication`) starts the Spring context and prints an encoded BCrypt example password to stdout (development helper). This indicates password encoding for login credentials is intended using `BCryptPasswordEncoder`.
- Spring Security is activated via `spring-boot-starter-security` and Thymeleaf is integrated with the `thymeleaf-extras-springsecurity5` module for security-aware templates.
- JPA/Hibernate is used for database access. Lazy loading is enabled for convenience via `spring.jpa.properties.hibernate.enable_lazy_load_no_trans=true`.
- JasperReports and JFreeChart libraries are used for report generation and charting.

## Configuration (important file: `src/main/resources/application.properties`)
Key values in repository (verify before using in another environment):

- spring.application.name = mfrs
- spring.datasource.url = jdbc:mysql://localhost:3306/dbmfrs
- spring.datasource.username = root
- spring.datasource.password = Viagbbi2003
- spring.datasource.driver-class-name = com.mysql.cj.jdbc.Driver
- spring.jpa.properties.hibernate.enable_lazy_load_no_trans = true
- spring.jpa.properties.hibernate.format_sql = true

Security/sensitive note: `application.properties` contains a plaintext database password in the repo; rotate or override in production (use environment variables or an externalized config mechanism).

## Database setup
1. Install MySQL (v8.x recommended) and create the database:
   - Database name expected: `dbmfrs`
2. Create or update the DB user with appropriate privileges for the schema.
3. Ensure `spring.datasource.*` properties in `application.properties` point to the production credentials or use environment variables to override at runtime.

Example SQL (adjust as needed):

```sql
CREATE DATABASE dbmfrs CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
-- create user and grant (optional if using root in dev)
CREATE USER 'mfrs_user'@'localhost' IDENTIFIED BY 'yourStrongPassword';
GRANT ALL ON dbmfrs.* TO 'mfrs_user'@'localhost';
FLUSH PRIVILEGES;
```

## Build, run and deploy
Prerequisites
- JDK 17 installed and JAVA_HOME set
- Maven 3.6+/wrapper included (use provided `mvnw.cmd` on Windows)
- MySQL server running and accessible

Options

1) Build the WAR using the Maven wrapper (Windows PowerShell):

```powershell
.\mvnw.cmd clean package -DskipTests
```

This produces `target/mfrs-0.0.1-SNAPSHOT.war`.

2) Deploy to Apache Tomcat (recommended because `tomcat` dependency is `provided`):
- Copy the WAR to `TOMCAT_HOME\webapps\` and start Tomcat. The application will be available at `http://<host>:<port>/<context>`.

3) Run from IDE / locally (development):
- If you prefer to run with embedded Tomcat using `spring-boot:run` or `java -jar`, remove or change the `provided` scope for `spring-boot-starter-tomcat` in `pom.xml` OR run via the IDE with the classpath including Tomcat jars. Alternatively, run from the IDE by launching `MfrsApplication` as a Java application.

To run with the Spring Boot plugin (may require adjusting `tomcat` scope):

```powershell
.\mvnw.cmd spring-boot:run
```

4) Running tests:

```powershell
.\mvnw.cmd test
```

## Reports and templates
- `company.jrxml` is the JasperReports source template used to generate PDF reports. The project includes `jasperreports` as a dependency and JFreeChart to emit charts inside reports.
- Templates under `src/main/resources/templates/reports` (and `templates/`) are used for web-based views and possibly for HTML report previews.

## Common developer tasks
- Regenerate compiled Jasper files (if the project compiles them at runtime, they may be loaded directly from the `.jrxml`): use the `jasperreports` APIs in code or precompile via build tasks.
- Change DB credentials: prefer `-Dspring.datasource.password=...` or environment variables, not committing secrets.
- Add a new Thymeleaf view: create `src/main/resources/templates/yourpage.html` and a controller with a `@GetMapping` returning the view name.
- Use `BCryptPasswordEncoder` for password hashing (example snippet is in `MfrsApplication`).

## Troubleshooting
- Common errors: DB connection refused — confirm MySQL is running and `application.properties` matches credentials.
- If you see ClassNotFound for servlet classes when running as a standalone JAR, the `provided` Tomcat scope is excluding embedded Tomcat. Either change scope or deploy as WAR to external container.
- View logs under Tomcat `logs/` or capture console output when running via `spring-boot:run`.

## Security & production hardening recommendations
- Never commit production secrets. Use environment variables, Spring Cloud Config, or vault solutions.
- Use a dedicated DB user with least privileges.
- Enable HTTPS at container/reverse-proxy level.
- Review Spring Security configuration (custom config classes) to ensure proper CSRF, session, and role-based controls.

## Next steps / recommended improvements
- Externalize secrets and configuration (example: use `SPRING_DATASOURCE_URL`, `SPRING_DATASOURCE_USERNAME`, `SPRING_DATASOURCE_PASSWORD` env vars).
- Add an automated CI pipeline to run `mvnw.cmd test` and build artifacts.
- Add startup instructions to include environment-specific profiles (Spring profiles) for dev/prod.

- Produce a detailed README describing technical specs, structure and how it works: Done.

