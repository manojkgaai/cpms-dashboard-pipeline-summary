import * as React from 'react';
import styles from './DashboardSummary.module.scss';
import { IDashboardSummaryProps } from './IDashboardSummaryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IDashboardSummaryState } from './IDashboardSummaryState';

export default class DashboardSummary extends React.Component<IDashboardSummaryProps, IDashboardSummaryState> {

  constructor(props: IDashboardSummaryProps, state: IDashboardSummaryState) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });

    this.getItems();
    this.state = { weekItems: [], totals: [] };
  }

  public getMonday(d) {
    d = new Date(d);
    var day = d.getDay(),
      diff = d.getDate() - day + (day == 0 ? -6 : 1); // sunday =0, adjust day from 0 to 6 in -6:_ to return day
    return new Date(d.setDate(diff));
  }

  public addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  public dateCheck(from, to, check) {

    var fDate, lDate, cDate;
    fDate = Date.parse(from);
    lDate = Date.parse(to);
    cDate = Date.parse(check);

    if ((cDate <= lDate && cDate >= fDate)) {
      return true;
    }
    return false;
  }


  public async getItems() {
    var today = new Date(new Date().setHours(0, 0, 0, 0));

    var curMonday = this.getMonday(today);
    var curTuesday = this.addDays(curMonday, 1);
    var curWednesday = this.addDays(curMonday, 2);
    var curThursday = this.addDays(curMonday, 3);
    var curFriday = this.addDays(curMonday, 4);
    var curSaturday = this.addDays(curMonday, 5);
    var curSunday = this.addDays(curMonday, 6);

    var weekmonday = [];
    var weektuesday = [];
    var weekwednesday = [];
    var weekthursday = [];
    var weekfriday = [];
    var weeksaturday = [];
    var weeksunday = [];
    var i = 0;
    var j = 0;
    for (i = 0; i < 8; i++) {
      weekmonday[i] = this.addDays(curMonday, j);
      weektuesday[i] = this.addDays(curTuesday, j);
      weekwednesday[i] = this.addDays(curWednesday, j);
      weekthursday[i] = this.addDays(curThursday, j);
      weekfriday[i] = this.addDays(curFriday, j);
      weeksaturday[i] = this.addDays(curSaturday, j);
      weeksunday[i] = this.addDays(curSunday, j);
      j += 7;
    }

    var PStotcndcount = 0;
    var PStotnpm = 0;
    var PStotwtst = 0;
    var PStotstarted = 0;
    var PStotnoinit = 0;
    var PStotinit = 0;
    var PStoteverified = 0;

    let obfield = this.props.onboardfield;
    let startdtfield = this.props.startdtfield;
    let npmfield = this.props.npmfield;

    let allItems: any[] = [];
    try{
      let query = "("+obfield + " ne 'Completed') and ("+obfield+" ne 'Backout')";
      allItems = await sp.web.lists.getById(this.props.list).items.filter(query).getAll();
      console.log(allItems);
    }
    catch(er){
      console.log("Error in SharePoint Rest Call");
      console.log(er);
    }

    let weekRows: any[] = [];
    let totalsRow: any[] = [];

    for (var a = 0; a < 8; a++) {

      var PSobcndcount = 0;
      var PSobnpm = 0;
      var PSwtst = 0;
      var PSstarted = 0;
      var PSnoinit = 0;
      var PSinit = 0;
      var PSeverified = 0;

      allItems.map(item => {

        var OBstatusval = item[obfield];
        var PSactstdate = item[startdtfield];
        var PSstdate = new Date(PSactstdate);
        var OBnpm = item[npmfield];

        if (this.dateCheck(weekmonday[a], weeksunday[a], PSstdate)) {
          PSobcndcount++;
          PSobnpm += OBnpm;

          if (OBstatusval == "WaitStart") {
            PSwtst++;
          }
          if (OBstatusval == "Started") {
            PSstarted++;
          }
          if (OBstatusval == "NoInit") {
            PSnoinit++;
          }
          if (OBstatusval == "Init") {
            PSinit++;
          }
          if (OBstatusval == "Everified") {
            PSeverified++;
          }
        }
      });

      PStotcndcount += PSobcndcount;
      PStotnpm += PSobnpm;
      PStotwtst += PSwtst;
      PStotstarted += PSstarted;
      PStotnoinit += PSnoinit;
      PStotinit += PSinit;
      PStoteverified += PSeverified;

      var weekstartdispdate = ('0' + weekmonday[a].getDate()).slice(-2);
      var weekstartdispmonth = ('0' + (weekmonday[a].getMonth() + 1)).slice(-2);
      var weekstartdispyear = weekmonday[a].getFullYear();
      var weekstartdispyearval = weekstartdispyear.toString();

      var weekenddispdate = ('0' + weeksunday[a].getDate()).slice(-2);
      var weekenddispmonth = ('0' + (weeksunday[a].getMonth() + 1)).slice(-2);
      var weekenddispyear = weeksunday[a].getFullYear();
      var weekenddispyearval = weekenddispyear.toString();

      var thisweek = weekstartdispmonth + "/" + weekstartdispdate + " - " + weekenddispmonth + "/" + weekenddispdate;
      var thisweekstartdt = weekstartdispmonth + "/" + weekstartdispdate + "/" + weekstartdispyearval;
      var thisweekenddt = weekenddispmonth + "/" + weekenddispdate + "/" + weekenddispyearval;
      var linkfilter = "("+this.props.startdtfield+" gt '"+thisweekstartdt+"') and ("+this.props.startdtfield+" lt '"+thisweekenddt+"')";
      console.log(linkfilter);

      var thisweekdata = PSobcndcount + PSobnpm + PSwtst + PSstarted + PSnoinit + PSinit + PSeverified;
      if (thisweekdata > 0) {
        weekRows.push({ "week": thisweek, "cndcount": PSobcndcount, "link": linkfilter, "npm": "$" + PSobnpm.toLocaleString("en-US"), "wtst": PSwtst, "started": PSstarted, "noinit": PSnoinit, "init": PSinit, "everified": PSeverified });
      }
    }

    totalsRow.push({ "week": "Total", "cndcount": PStotcndcount, "npm": "$" + PStotnpm.toLocaleString("en-US"), "wtst": PStotwtst, "started": PStotstarted, "noinit": PStotnoinit, "init": PStotinit, "everified": PStoteverified });

    this.setState({ weekItems: weekRows, totals: totalsRow });
  }

  public renderTableData() {
    return this.state.weekItems.map((item, index) => {
      const { week, cndcount, link, npm, wtst, started, noinit, init, everified } = item;
      return (
        <div className={styles.divRow}>
          <div className={styles.divLabel}>{week}</div>
          <div className={styles.divCell}><a href={this.props.linkPageUrl + "?filterQuery="+link} style={{cursor: "pointer", textDecoration:"none" }} >{cndcount}</a></div>
          <div className={styles.divCell}>{npm}</div>
          <div className={styles.divCell}>{wtst}</div>
          <div className={styles.divCell}>{started}</div>
          <div className={styles.divCell}>{noinit}</div>
          <div className={styles.divCell}>{init}</div>
          <div className={styles.divCell}>{everified}</div>
        </div>
      );
    });
  }

  public renderFooterData() {
    return this.state.totals.map((item, index) => {
      const { week, cndcount, npm, wtst, started, noinit, init, everified } = item;
      return (
        <div className={styles.footRow}>
          <div className={styles.divLabel}>{week}</div>
          <div className={styles.divCell}>{cndcount}</div>
          <div className={styles.divCell}>{npm}</div>
          <div className={styles.divCell}>{wtst}</div>
          <div className={styles.divCell}>{started}</div>
          <div className={styles.divCell}>{noinit}</div>
          <div className={styles.divCell}>{init}</div>
          <div className={styles.divCell}>{everified}</div>
        </div>
      );
    });
  }

  public render(): React.ReactElement<IDashboardSummaryProps> {

    return (
      <div className={styles.dashboardSummary}>
        <div>
          <div className={styles.wpTitle}>{this.props.wptitle}
          </div>
        </div>
        <div className={styles.divTable}>
          <div className={styles.headRow}>
            <div className={styles.divLabel}>Week</div>
            <div className={styles.divCell}>Start</div>
            <div className={styles.divCell}>NPM</div>
            <div className={styles.divCell}>WaitStart</div>
            <div className={styles.divCell}>Started</div>
            <div className={styles.divCell}>NoInit</div>
            <div className={styles.divCell}>Init</div>
            <div className={styles.divCell}>Everified</div>
          </div>
          {this.renderTableData()}
          {this.renderFooterData()}
        </div>
      </div>
    );
  }
}
