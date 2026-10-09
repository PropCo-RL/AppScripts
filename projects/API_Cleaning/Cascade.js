/**
 * Master Controller: Full GCal Sync Cycle
 * Runs all scripts in the required order.
 */
function Cascade() {

  console.log("--- START: GCal Sync Cycle ---");

  try {

    // STEP 1
    console.log("STEP 1: Import current Calendar...");
    importCalendarToGCal();

    Utilities.sleep(3000);

    // STEP 2
    console.log("STEP 2: Generate initial Audit...");
    generateGCalAudit();

    Utilities.sleep(3000);

    // STEP 3
    console.log("STEP 3: Apply vacation logic...");
    applyVacationLogicToGCal();

    Utilities.sleep(3000);

    // STEP 4
    console.log("STEP 4: Sync changes to Calendar...");
    syncGCalWithAudit();

    // VERY IMPORTANT
    // Give Google Calendar time to process creates/deletes
    Utilities.sleep(10000);

    // STEP 5
    console.log("STEP 5: Refresh Calendar import...");
    importCalendarToGCal();

    Utilities.sleep(5000);

    // STEP 6
    console.log("STEP 6: Finalize Audit...");
    generateGCalAudit();

    console.log("--- SUCCESS: Full GCal Sync Complete ---");

  } catch (e) {

    console.error("MASTER ERROR: " + e.toString());

  }

}
