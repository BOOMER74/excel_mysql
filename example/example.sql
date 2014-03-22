CREATE DATABASE IF NOT EXISTS `excel_mysql_base`;

USE `excel_mysql_base`;

CREATE TABLE IF NOT EXISTS `excel_mysql` (
	`column0` TEXT NOT NULL,
	`column1` TEXT NOT NULL,
	`column2` TEXT NOT NULL
);

INSERT INTO `excel_mysql` (`column0`, `column1`, `column2`)
VALUES ('1', '2', '3'), ('4', '5', '6'), ('7', '8', '9');