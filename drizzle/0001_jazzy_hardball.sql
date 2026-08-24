CREATE TABLE `reconciliation_reviews` (
	`id` text PRIMARY KEY NOT NULL,
	`reconciliation_id` text NOT NULL,
	`row_key` text NOT NULL,
	`row_json` text NOT NULL,
	`decision` text NOT NULL,
	`note` text NOT NULL,
	`proposed_entry` text NOT NULL,
	`reviewer` text NOT NULL,
	`created_at` integer NOT NULL
);
