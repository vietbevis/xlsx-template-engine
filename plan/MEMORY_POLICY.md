# Memory Policy

## Goal

Moi phase sau khong can doc lai toan bo lich su. Sau khi hoan thanh mot phase, ghi mot note ngan vao memory de luu nhung gi da lam, file da sua, API da them va test da chay.

## When To Write

Ghi memory note sau moi phase hoac sau mot cum thay doi lon co the anh huong phase sau.

Khong ghi lai toan bo noi dung plan. Chi ghi quyet dinh va ket qua thuc te.

## Note Format

```md
# xlsx-template-engine phase N

Completed:
- ...

Files changed:
- ...

Public API added/changed:
- ...

Validation/tests:
- ...

Open decisions:
- ...
```

## What To Include

- Phase da hoan thanh.
- Files/folders quan trong da tao hoac sua.
- Public API moi hoac thay doi compatibility.
- Test/command da chay va ket qua.
- Quyet dinh can nho cho phase sau.
- Viec chua lam co chu y.

## What To Avoid

- Khong copy lai nguyen plan phase.
- Khong luu output dai cua command neu khong can.
- Khong luu secret, path nhay cam, token, credential.
- Khong ghi chung chung; memory phai giup phase sau bat dau nhanh.

