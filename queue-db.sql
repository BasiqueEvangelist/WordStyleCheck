CREATE DATABASE IF NOT EXISTS `wordstylecheck`;

CREATE TABLE IF NOT EXISTS `wordstylecheck`.`bl_task` (
    `TaskId` int(11) PRIMARY KEY AUTO_INCREMENT,
    `TaskType` varchar(45) NOT NULL,
    `TaskData` text NOT NULL,
    `CompletionStatus` tinyint(4) unsigned NOT NULL DEFAULT 1,
    `ResultData` text DEFAULT NULL,
    `Comment` text DEFAULT NULL,
    `CreateTs` timestamp NOT NULL DEFAULT current_timestamp(),
    `UpdateTs` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
);

CREATE TABLE IF NOT EXISTS `wordstylecheck`.`bl_pendingtask` (
    `PendingTaskId` int(11) PRIMARY KEY AUTO_INCREMENT,
    `TaskId` int(11) NOT NULL REFERENCES `bl_task`(`TaskId`),
    `ProcessingStatus` tinyint(3) unsigned NOT NULL DEFAULT 1,
    `ExpireTs` timestamp NOT NULL,
    `CreateTs` timestamp NOT NULL DEFAULT current_timestamp(),
    `UpdateTs` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
);