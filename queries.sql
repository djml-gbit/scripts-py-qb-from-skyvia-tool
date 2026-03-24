-- Facturas de proveedor
SELECT *
FROM Bill b
LEFT OUTER JOIN BillLineItem li
  ON b.Id = li.BillId
ORDER BY substr(b.TxnDate, 1, 10), b.DocNumber;

-- Bill payments
SELECT *
FROM BillPayment bp
LEFT OUTER JOIN BillPaymentLineItem li
  ON bp.Id = li.BillPaymentId
ORDER BY substr(bp.TxnDate, 1, 10), bp.DocNumber;

-- Diarios
SELECT *
FROM JournalEntry je
LEFT OUTER JOIN JournalEntryLineItem li
  ON je.Id = li.JournalEntryId
ORDER BY substr(je.TxnDate, 1, 10), je.DocNumber;


-- Notas de credito
SELECT *
FROM CreditMemo c
LEFT OUTER JOIN CreditMemoLineItem li
  ON c.Id = li.CreditMemoId
ORDER BY substr(c.TxnDate, 1, 10), c.DocNumber, li.LineNum;

-- Pagos de clientes
SELECT
  pli.PaymentId,
  p.Id AS PaymentId2,
  p.TxnDate AS PaymentDate,
  p.CustomerRefId,
  p.CustomerRefName,
  p.TotalAmt AS PaymentTotal,
  pli.Amount AS AmountApplied,
  pli.LinkedTxn
FROM PaymentLineItem pli
JOIN Payment p ON p.Id = pli.PaymentId
WHERE p.TxnDate >= '2025-12-01'
  AND p.TxnDate <= '2026-02-28'
ORDER BY p.TxnDate, pli.PaymentId;

-- Facturas de clientes
SELECT *
FROM Invoice i
LEFT OUTER JOIN InvoiceLineItem li
  ON i.Id = li.InvoiceId
WHERE substr(i.TxnDate, 1, 10) >= '2026-01-01'
  AND substr(i.TxnDate, 1, 10) <= '2026-01-31'
ORDER BY substr(i.TxnDate, 1, 10), i.DocNumber, li.LineNum;