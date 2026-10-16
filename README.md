# RFM format converter

Converts a wide price workbook (one row per restaurant, an In-Store / Delivery
column pair per item) into the long RFM upload format:

| Store ID | Item ID | In-Store Recommended Price | Delivery Recommended Price |
| --- | --- | --- | --- |

## Web page

`index.html` is the whole tool. It runs in the browser, and the workbook is never uploaded.

To host it on GitHub Pages:

1. Push this folder to a GitHub repository.
2. In the repository, open **Settings > Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, pick your branch and the `/ (root)` folder, then save.
4. After a minute the page is live at `https://<your-username>.github.io/<repo-name>/`.

The page loads SheetJS (xlsx 0.18.5) from cdnjs. To work offline or avoid the CDN,
download `xlsx.full.min.js` into the repo and change the `<script src>` to point to it.

## Command line

```
pip install openpyxl
python convert_to_rfm.py input.xlsx
python convert_to_rfm.py input.xlsx -o output_RFM_Format.xlsx
```

## Expected input layout

- A row of item headers such as `#10` + newline + `McRib`, merged across two columns
- Below it, a header row with `NSN`, then `In-Store` and `Delivery` under each item
- One restaurant per row, with its NSN as the Store ID

Any number of items and restaurants is supported.
