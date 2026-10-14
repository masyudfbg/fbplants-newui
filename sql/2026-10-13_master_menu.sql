-- =====================================================================
-- Menu "Master Menu" di SYS-SYSTEM (halaman MasterMenu.asp)
-- Aman dijalankan berulang (hanya menambah bila belum ada).
-- Akses hanya untuk role ARJ_IT.
-- Batalkan:  DELETE FROM menu_roles_tab WHERE fcmenucode = '1450000';
--            DELETE FROM menu_tab WHERE fccode = '1450000';  COMMIT;
-- =====================================================================
INSERT INTO menu_tab (fccode, fcname, menuparentcode, formlink, module, isactive, helpdescription,
                      sequence, remarks, fcentry, fcedit, fcip, lastupdate, lasttime, menutypecode)
SELECT '1450000', 'Master Menu', '1550000', 'MasterMenu', 14, 'TRUE', 'Aktif/nonaktifkan menu aplikasi',
       10, 'Mengatur ISACTIVE MENU_TAB', 'system', 'system', 'localhost', SYSDATE, TO_CHAR(SYSDATE, 'HH24:MI'), 'FORM'
  FROM dual
 WHERE NOT EXISTS (SELECT 1 FROM menu_tab WHERE fccode = '1450000');

INSERT INTO menu_roles_tab (fcrolescode, fcmenucode, fcentry, fcedit, fcip, lastupdate, lasttime)
SELECT 'ARJ_IT', '1450000', 'system', 'system', 'localhost', SYSDATE, TO_CHAR(SYSDATE, 'HH24:MI')
  FROM dual
 WHERE NOT EXISTS (SELECT 1 FROM menu_roles_tab x WHERE x.fcrolescode = 'ARJ_IT' AND x.fcmenucode = '1450000');

-- hanya ARJ_IT: cabut role lain bila skrip versi lama sempat dijalankan
DELETE FROM menu_roles_tab WHERE fcmenucode = '1450000' AND fcrolescode <> 'ARJ_IT';

COMMIT;
