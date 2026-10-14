-- =====================================================================
-- Antrian email PO yang sudah disetujui approver terakhir.
-- Jalankan sekali di skema aplikasi FB-Plants (DEV dulu, lalu PROD).
-- Dipakai oleh: email/po_mail_queue.inc (INSERT) dan
--               email/po_mail_sender.inc (proses kirim).
-- =====================================================================

CREATE TABLE PO_EMAIL_OUTBOX (
    ID           NUMBER          NOT NULL,
    FCBA         VARCHAR2(20)    NOT NULL,
    POCODE       VARCHAR2(60)    NOT NULL,
    STATUS       VARCHAR2(10)    DEFAULT 'PENDING' NOT NULL,  -- PENDING / SENDING / SENT / FAILED
    ATTEMPTS     NUMBER(3)       DEFAULT 0 NOT NULL,
    NEXT_TRY     DATE,                                         -- jadwal percobaan ulang setelah gagal
    LAST_ERROR   VARCHAR2(2000),
    MAIL_TO      VARCHAR2(500),                                -- penerima sebenarnya saat terkirim
    MAIL_CC      VARCHAR2(500),
    IS_TEST      CHAR(1),                                      -- Y = terkirim dalam mode uji coba
    PDF_FILE     VARCHAR2(200),                                -- nama file di email/po_pdf/
    FCENTRY      VARCHAR2(50),                                 -- user yang melakukan approval final
    FCIP         VARCHAR2(50),
    CREATED_AT   DATE            DEFAULT SYSDATE NOT NULL,
    UPDATED_AT   DATE,
    SENT_AT      DATE,
    CONSTRAINT PO_EMAIL_OUTBOX_PK PRIMARY KEY (ID),
    CONSTRAINT PO_EMAIL_OUTBOX_ST CHECK (STATUS IN ('PENDING', 'SENDING', 'SENT', 'FAILED'))
);

CREATE INDEX PO_EMAIL_OUTBOX_IX1 ON PO_EMAIL_OUTBOX (STATUS, NEXT_TRY);
CREATE INDEX PO_EMAIL_OUTBOX_IX2 ON PO_EMAIL_OUTBOX (FCBA, POCODE);

CREATE SEQUENCE PO_EMAIL_OUTBOX_SEQ START WITH 1 INCREMENT BY 1 NOCACHE;

-- Kirim ulang email yang gagal (mis. setelah App Password diperbaiki):
--   UPDATE PO_EMAIL_OUTBOX SET STATUS = 'PENDING', ATTEMPTS = 0, NEXT_TRY = NULL WHERE STATUS = 'FAILED';
--   COMMIT;
-- Riwayat pengiriman satu PO:
--   SELECT * FROM PO_EMAIL_OUTBOX WHERE POCODE = '...' ORDER BY ID DESC;
