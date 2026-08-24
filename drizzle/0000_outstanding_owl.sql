CREATE TABLE `bank_reconciliations` (
	`id` text PRIMARY KEY NOT NULL,
	`bank_name` text NOT NULL,
	`account_name` text NOT NULL,
	`period` text NOT NULL,
	`result_json` text NOT NULL,
	`created_at` integer NOT NULL
);
