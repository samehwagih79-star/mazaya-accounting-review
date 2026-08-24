import {integer,sqliteTable,text} from "drizzle-orm/sqlite-core";

export const bankReconciliations=sqliteTable("bank_reconciliations",{
  id:text("id").primaryKey(),
  bankName:text("bank_name").notNull(),
  accountName:text("account_name").notNull(),
  period:text("period").notNull(),
  resultJson:text("result_json").notNull(),
  createdAt:integer("created_at").notNull(),
});

export const reconciliationReviews=sqliteTable("reconciliation_reviews",{
  id:text("id").primaryKey(),
  reconciliationId:text("reconciliation_id").notNull(),
  rowKey:text("row_key").notNull(),
  rowJson:text("row_json").notNull(),
  decision:text("decision").notNull(),
  note:text("note").notNull(),
  proposedEntry:text("proposed_entry").notNull(),
  reviewer:text("reviewer").notNull(),
  createdAt:integer("created_at").notNull(),
});
