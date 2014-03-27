CREATE DATABASE IF NOT EXISTS `excel_mysql_base`;

USE `excel_mysql_base`;

CREATE TABLE IF NOT EXISTS `excel_mysql_data` (
	`id`         INT(11)      NOT NULL AUTO_INCREMENT,
	`first_name` VARCHAR(50)  NOT NULL,
	`last_name`  VARCHAR(50)  NOT NULL,
	`email`      VARCHAR(100) NOT NULL,
	`pay`        FLOAT(10, 2) NOT NULL,
	PRIMARY KEY (`id`)
);

INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`)
VALUES ('John', 'Smith', 'j.smith@example.com', 10000.00);

INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`)
VALUES ('Steve', 'Smith', 's.smith@example.com', 11000.00);

INSERT INTO `excel_mysql_data` (`first_name`, `last_name`, `email`, `pay`)
VALUES ('Oscar', 'Wild', 'o.wild@example.com', 12250.59);