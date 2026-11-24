# CarWatch group

The group keeps ads from sites CarWatch can't read (OLX, Mercado Livre, Webmotors, Facebook Marketplace...) in a
Google Sheet that you own. Everyone you invite sees them in their CarWatch list, with a score and a verdict; the group
code lets them add ads. It's free, and the people you invite need no account.

## Set it up (once, about 5 minutes; easiest on a computer)

1. Create a new Google Sheet (open `sheets.new`) and name it **CarWatch group**.
2. In the sheet: **Extensions > Apps Script**. Delete the sample code, paste all of [Code.gs](Code.gs) and save
   (the disk icon).
3. In the function menu next to **Run**, pick `setup` and press **Run**. Google asks for permission: **Review
   permissions**, pick your account, then **Advanced > Go to … (unsafe)** and **Allow**. The script can only read and
   write this one spreadsheet.
4. Open the **Execution log** and copy the **group code**. The sheet now has an **Ads** tab.
5. **Deploy > New deployment**. Click the gear next to "Select type" and pick **Web app**. Set **Execute as: Me** and
   **Who has access: Anyone**, then **Deploy**, and copy the **Web app URL** (it ends in `/exec`).
6. On your phone, in CarWatch: **Menu (⋯) > Group**. Paste the address and the code and tap **Save**. Then
   **Invite someone** sends a link that joins the group in one tap.

## Use it

- **Android:** install CarWatch (Chrome menu ⋮ > Install app, or Add to home screen). Then share an ad from the OLX,
  Mercado Livre, Webmotors or Facebook app, or from Chrome, and pick **CarWatch**. It reads the link and the text,
  fills in the model, year, km, price, version and town it finds, and shows the verdict. Fix what it got wrong, then
  **Send to the group**, or **Save to my list** to keep it on your phone only.
- **iPhone or computer:** Menu > Check an ad, and paste the link or the text.
- In the list, **Saved and group** shows only those ads. **Sold or removed**, in an ad's details, takes it out for
  everyone. An ad nobody shares again for 60 days drops out by itself.

## Good to know

- Anyone with the invite link can see the group's ads; only people with the code can add them or mark them sold.
  To change the code, run `newCode()` in the editor (the new code is in the log) and send new invites.
- The sheet keeps only public listing data: link, site, model, year, km, price, version, gearbox, body, town and the
  ad's title. Open it to see or clean up the ads.
- After editing Code.gs, publish it with **Deploy > Manage deployments**, the pencil, **Version: New version**,
  **Deploy**, so the address stays the same.
- Shared ads get the same fair price and score as the cars CarWatch reads, but they stay out of the market
  statistics in the Prices tab, since they're typed or read from an ad's text.
