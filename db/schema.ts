import {integer,sqliteTable,text} from "drizzle-orm/sqlite-core";

export const bankReconciliations=sqliteTable("bank_reconciliations",{
  id:text("id").primaryKey(),
  bankName:text("bank_name").notNull(),
  accountName:text("account_name").notNull(),
  period:text("period").notNull(),
  resultJson:text("result_json").notNull(),
  createdAt:integer("created_at").notNull(),
});
