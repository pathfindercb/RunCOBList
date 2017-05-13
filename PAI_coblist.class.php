<?php
/** Updated 20170513
 * TODO - finish slip dupes, email logrun 
 * package    PAI_COBList
 * @license        Copyright © 2017 Pathfinder Associates, Inc.
 */
class COBList
{
    /**
     * Retrieves the CSV file exported by Manage Users and creates the Excel XML COBList
     */   
	// Private Variables //
		const iVersion = "3.0.1";
		private $dbUser = array();
		private $hdrUser = array();
		private $dbRes = array();
		private $hdrRes = array();
		private $dbSold = array();
		private $hdrSold = array();
		private $dbGone = array();
		private $hdrGone = array();
		private $dbErr = array();
		private $hdrErr = array();
		private $dbSlip = array();
		private $hdrSlip = array();
		private $hdrWait = array();
		private $dbKayak = array();
		private $hdrKayak = array();
		private $dbRenter = array();
		private $hdrRenter = array();
		private $dbStaff = array();
		private $hdrStaff = array();
		private $dbGridT1 = array();
		private $hdrGridT1 = array();
		private $dbGridT2 = array();
		private $hdrGridT2 = array();
		private $dbVoterT1 = array();
		private $hdrVoterT1 = array();
		private $dbVoterT2 = array();
		private $hdrVoterT2 = array();
		private $dbPets = array();
		private $hdrPets = array();
		private $dbVoter = array();
		private $hdrVoter = array();
		private $pdo;

	//Constants for fields in User tab
		const iCreated = 0;
		const iUsername = 1;
		const iEnabled = 2;
		const iFirstName = 3;
		const iEmail = 4;
		const iAccess = 5;
		const iUser1LastName = 6;
		const iUnit = 7;
		const iHomePhone = 8;
		const iUser1WorkPhone = 9;
		const iUser1CellPhone = 10;
		const iUser1Occupation = 11;
		const iUser1Employer = 12;
		const iUser1Hobbies = 13;
		const iUser2FirstName = 14;
		const iUser2LastName = 15;
		const iUser2Email = 16;
		const iUser2WorkPhone = 17;
		const iUser2CellPhone = 18;
		const iUser2Occupation = 19;
		const iUser2Employer = 20;
		const iUser2Hobbies = 21;
		const iChild1Name = 22;
		const iChild2Name = 23;
		const iChild3Name = 24;
		const iChild4Name = 25;
		const iChild1Birthdate = 26;
		const iChild2Birthdate = 27;
		const iChild3Birthdate = 28;
		const iChild4Birthdate = 29;
		const iOwner = 30;
		const iMailings = 31;
		const i2ndAddress = 32;
		const iEmergencyContact = 33;
		const iUnitWatcher = 34;
		const iStack = 35;
		const iSlip = 36;
		const iPets = 37;
		const iOfficialVoter = 38;
		const iShowProfile = 39;
		const iShowEmail = 40;
		const iShowPhone = 41;
		const iShowChildren = 42;
		const iAdminNotes = 43;
		const iLastLogin = 44;
		const iUserID = 45;
		const iVoter = 46;
		const iAddress = 47;
		const iCityStateZip = 48;
		const iFloor = 49;
	
	// indicate whether to show all phone/email or use Profile for external use
	public $showInfo = true;
	// indicate whether full export so trigger all error messages
	public $fullRun = true;
		

	public function __construct ()
	{
		date_default_timezone_set('America/New_York');
	}
	
	public function Checkfile(&$checkmsg)
    {
	// called by loadcsv to upload the file and check for format/size & load to dbUser
		$checkmsg="";
		if($_FILES["import"]["error"] > 0){
			$checkmsg =  "Error: " . $_FILES["import"]["error"] . "<br>";
			$checkmsg .=  "Error on upload. Please click the left menu item to re-run";
			return false;
		} 
		
		// get file info
		$filename = $_FILES["import"]["name"];
		$filetype = $_FILES["import"]["type"];
		$filesize = $_FILES["import"]["size"];
		$tempfile = $_FILES["import"]["tmp_name"];
		$maxsize = 1 * 1024 * 1024;
	
		// Verify file extension
		$ext = pathinfo($filename, PATHINFO_EXTENSION);
		if(!($ext == "csv")) {
			$checkmsg =  "Error: Please select a CSV file format.";
			return false;
		} 
		// Verify file size - 1MB maximum
		if($filesize > $maxsize) {
			$checkmsg =  "Error: Please select a smaller CSV file.";
			return false;
		} 
		$temp = file($tempfile);
		
		// convert to 2 dimensional
		foreach ($temp as $line) {
			$this->dbUser[] = str_getcsv($line);
		}		

		// check if exported from website
		if (!(trim($this->dbUser[0][0])=="Created")) {
			$checkmsg =  "Error: '" . strlen($this->dbUser[0][0]) . "' Not valid exported CSV file.";
			return false;
		}
		if (!(count($this->dbUser[0])==45)) {
			$checkmsg =  "Error: " . count($this->dbUser[0]) . " Not valid exported CSV file.";
			return false;
		}

		return true;
	}
	
	public function ProcessFile(&$checkmsg)
	{
		// Called by loadcsv to process dbUser and build all worksheets
		//open database so the pdo object is available to all functions
		if (! $this->opendb($checkmsg)) {
			return false;
		}
		if ($this->fullRun) {
			// delete all records from Slips & WaitList table and reset owner count in UnitMaster
			$stmt = $this->pdo->query('DELETE FROM Slips');
			$stmt = $this->pdo->query('DELETE FROM WaitList');
			$stmt = $this->pdo->query('UPDATE UnitMaster SET count = 0, Owner = "", Renter = "", Voter = "", Address = "", CityStateZip = ""') ;
		}
		// remove header 
		$this->hdrUser = array_shift($this->dbUser);

		// extend header for userid, voter, addr, city, floor, lease start,end
		$this->hdrUser[] = "UserID";
		$this->hdrUser[] = "Voter";
		$this->hdrUser[] = "Address";
		$this->hdrUser[] = "CityStateZip";
		$this->hdrUser[] = "Floor";
		//header flipped set to strings
		$keys = array_keys(array_flip($this->hdrUser));
		$this->hdrUser = array_fill_keys($keys, "string");
		// now build proper address fields, floor, etc.
		$this->BuildAddress();
		$this->BuildVoter();
		$this->CheckData();
		if ($this->fullRun) {
			// now checkdata for errors
			$this->BuildListing();
			$this->BuildSlip();
			$this->BuildStaff();
			$this->BuildErr();
			$this->BuildGrids();
		}
			// now create Excel file
		$this->CreateFile();
		//now log this run
		$this->LogRun();

		// end of processfile
		unset($this->pdo);

		return true;
	}
	
	function CheckData() {
	//test for errors
	//Unit format wrong, Stack mismatch, Access mismatch, Phone format, 2nd Address format, Emergency contact phone, Owner access, Mailing blank
	//Offical Voter missing or duplicate

	//step thru User database
	foreach ($this->dbUser as $rowData) {
		if ((stripos($rowData[self::iUnit],'gone') !== false)) {
		} elseif (stripos($rowData[self::iUnit],'sold') !== false) {
		} else {
			// check if not Enabled
			if(($rowData[self::iEnabled] == "No") && stripos($rowData[self::iEmail],0, 11) !== "holdall@gmx" ){ 
					$this->addError('8','Not enabled',$rowData[self::iUnit],$rowData[self::iUser1LastName],'Has email but not enabled');
			} 
			// check 2nd address format
			if (count(explode(',',$rowData[self::i2ndAddress]))!==2 && strlen($rowData[self::i2ndAddress]) > 0 ){
				$this->addError('10','2ndAddress format',$rowData[self::iUnit],$rowData[self::iUser1LastName],$rowData[self::i2ndAddress]);
			}
			//check missing Emergency Contact
			if ($rowData[self::iEmergencyContact] == "") {
				$this->addError('14','Emergency contact',$rowData[self::iUnit],$rowData[self::iUser1LastName],'No Emergency Contact');
			}
			//check missing email address
			if ($rowData[self::iEmail] == "" ){
				$this->addError('2','Email address',$rowData[self::iUnit],$rowData[self::iUser1LastName],'No email address');
			}
			//check phone format
			if(!$this->CheckPhoneFormat($rowData[self::iHomePhone]))
			{
				$this->addError('12','Phone Format - Home',$rowData[self::iUnit],$rowData[self::iUser1LastName],$rowData[self::iHomePhone]);
			}
			if(!$this->CheckPhoneFormat($rowData[self::iUser1CellPhone]))
			{
				$this->addError('12','Phone Format - Cell',$rowData[self::iUnit],$rowData[self::iUser1LastName],$rowData[self::iUser1CellPhone]);
			}
			if(!$this->CheckPhoneFormat($rowData[self::iUser1WorkPhone]))
			{
				$this->addError('12','Phone Format - Work',$rowData[self::iUnit],$rowData[self::iUser1LastName],$rowData[self::iUser1WorkPhone]);
			}
			
			//check unit format
			$this->CheckUnitFormat($rowData);
		}
	} 
	//check owner count in UnitMaster
	$query1 = $this->pdo->prepare("SELECT Unit, count, Owner FROM UnitMaster WHERE count > 1");
	$query1->execute();
	while ($row = $query1->fetch(PDO::FETCH_ASSOC)) {
		$this->addError('16','Owner Count',$row['Unit'],$row['count'],$row['Owner']);
		}
	//check missing Voter in UnitMaster
	$query1 = $this->pdo->prepare("SELECT Unit, Owner, Voter FROM UnitMaster WHERE Voter = ''");
	$query1->execute();
	while ($row = $query1->fetch(PDO::FETCH_ASSOC)) {
		$this->addError('13','Missing Voter',$row['Unit'],$row['Owner'],'Check Official Voter');
		}
	//now query to dbVoter
	$query1 = $this->pdo->prepare("SELECT Voter, Address, CityStateZip, Unit, Bldg
						FROM UnitMaster 
						ORDER BY Unit");
	$query1->execute();
	$this->dbVoter = $query1->fetchALL(PDO::FETCH_ASSOC);

	}
	
	function BuildAddress()
	{
	// loop thru db and build calculated fields like addr, voter, floor
	// also build Resident Listing, Staff, Sold, Gone, and Renter arrays
	foreach ($this->dbUser as &$rowData) {
	//skip if gone/sold or staff
		if (stripos($rowData[self::iUnit],'gone') !== false) {
			//build gone db
			$this->dbGone[]=$rowData;
		} elseif (stripos($rowData[self::iUnit],'sold') !== false) {
			//build sold db
			$this->dbSold[]=$rowData;
		} elseif
			(((stripos($rowData[self::iAccess],'^A') !== false) &&
			(stripos($rowData[self::iAccess],'^ADMINWM') == false))||
			($rowData[self::iAccess]=='ADMINEM'))
			{
			// build staff db
			//write row - 6,3,4,5
			$this->dbStaff[]=array($rowData[6],$rowData[3],$rowData[4],$rowData[5]);
		} else {
			// first get voter for this record and add a column
			$rowData[]= $this->GetVoter($rowData);
			
			// then build addr, citystate for mailings based on user settings
			$temp = $this->GetAddress($rowData);
			$rowData[] = $temp[0];
			$rowData[] = $temp[1];
			
			// then get floor for all units in this row
			$rowData[] = $this->GetFloor($rowData[self::iUnit]);
			
			// copy current resident to dbRes
			//write row - 6,3,7,8,10,4,33,34
			$this->dbRes[]=array(
				$rowData[6],$rowData[3],$rowData[7],$rowData[8],
				$rowData[10],$rowData[4],$rowData[33],$rowData[34]
			);
		
			// copy certain fields if current renter to dbRenter
			if ($rowData[self::iAccess] == "MEMBER") {
				$temp = $this->GetLeaseDates($rowData);
				//write row - GetLeaseDates(38),6,3,8,10,4,33,34
				$this->dbRenter[]=array(
					$temp[1], $temp[0],
					$rowData[7],$rowData[6],$rowData[3],$rowData[8],
					$rowData[10],$rowData[4],$rowData[33],$rowData[34]
				);
				if ($rowData[self::iOwner] == "Yes") {
					$this->addError('1','Owner Error',$rowData[self::iUnit],$rowData[self::iUser1LastName],'Owner with only member access');
				}
			}	
			// copy certain fields to dbPets
			if (strlen(trim($rowData[self::iPets])) > 0) {
				//write row - 7,6,3,37,8,10,4,33,34
				$this->dbPets[]=array(
					$rowData[7],$rowData[6],$rowData[3],$rowData[37], $rowData[8],
					$rowData[10],$rowData[4],$rowData[33],$rowData[34]
				);
			}
		}
	}
	return;
	}

	function BuildVoter ()
	{
		$this->hdrVoter = array('Name'=>'string', 'Address'=>'string','CityStateZip'=>'string','Unit'=>'string','Bldg'=>'string');

	}
	function BuildListing()
	{
		//setup Listing columns for dbRes & dbRenter that was build in BuildAddress
		$this->hdrRes = array('Last Name'=>'string', 'First Name'=>'string','Unit'=>'string','Home Phone'=>'string','Cell Phone'=>'string','Email'=>'string','Emergency Contact'=>'string','Unit Watcher'=>'string');
		$this->hdrRenter = array('Lease End'=>'string', 'Lease Start'=>'string', 'Unit'=>'string','Last Name'=>'string', 'First Name'=>'string','Home Phone'=>'string','Cell Phone'=>'string','Email'=>'string','Emergency Contact'=>'string','Unit Watcher'=>'string');
		$this->hdrPets = array('Unit'=>'string','Last Name'=>'string', 'First Name'=>'string','Pets & WSD/ESA'=>'string','Home Phone'=>'string','Cell Phone'=>'string','Email'=>'string','Emergency Contact'=>'string','Unit Watcher'=>'string');
		return;
	}

	function BuildStaff()
	{
		//setup columns for dbStaff that was build in BuildAddress
		$this->hdrStaff=array('Last Name'=>'string', 'First Name'=>'string','Email'=>'string','Access'=>'string');
		return;
	}

	function BuildErr()
	{
		//setup Error columns for dbErr
		$this->hdrErr = array('Level'=>'string', 'Function'=>'string','Unit'=>'string','Name'=>'string','Message'=>'string');
		return;
	}

	function BuildSlip() {
		// scan dbUser to build Slips and Kayak db and arrays
//NEED TO MERGE TESTSLIP TO CHECK FOR MULTI UNITS ASSIGNED TO A SLIP
		// build header
		$temp = ($this->showInfo) ? 'Internal':'External';
		$this->hdrSlip = array('Dock'=>'string', 'Slip'=>'string','Class'=>'string','Rate'=>'string','Type'=>'string','Condition'=>'string', 'Last Name'=>'string','Unit'=>'string','Lift'=>'string','Phone'=>'string','Email'=>'string');
		$this->hdrWait = array('Date'=>'string','Name'=>'string','Unit'=>'string','Number'=>'string');
		
		// step thru each line of the file
		foreach ($this->dbUser as $row) {
			// skip if gone or sold
			if (stripos($row[7],"GONE")!== false) {
				continue;
			}
			if (stripos($row[7],"SOLD")!== false) {
				continue;
			}
			// if no slip then skip
			if (empty($row[36])){
				continue;
			}
			// explode multiple slips
			$temp = explode (';',$row[36]);
			foreach ($temp as $slip) {
				// check if waitlist and skip
				if (stripos($slip,"W")!== false) {
					// now add to WaitList table
					$wdate = date("Y.m.d",strtotime(substr(trim($slip),3,8)));
					$sql = "INSERT INTO WaitList (type,unit, names, date, number)
							VALUES ('" . substr(trim($slip),1,1) . "','" . $row[7] . "', '" . $row[6] . "', '" . $wdate . "', '" . substr(trim($slip),9,1) . "')";
							// execute the SQL statement - if returns fail then report
					if ($this->pdo->query($sql)){
					} else {
						echo "Failed " . $slip . "<br>";
					}
					
				} else {
					// if slip has L then strip off L and set lift  = true
					if (stripos($slip,"L")){
						$lift = 1;
						$slip = substr(trim($slip),0,3);
					} else {
						$lift = 0;
					}
					// decide if include email and phone based on user Profile settings if showInfo property is set = false
					$phone = "";
					$email = "";
					if ($this->showInfo) {
						$phone = $this->GetBestPhone($row);
						$email = $row[self::iEmail];
					} elseif ($row[self::iShowProfile]=='Yes') {
							if($row[self::iShowPhone]=='Yes') {
								$phone = $this->GetBestPhone($row);
							}
							if($row[self::iShowEmail]=='Yes') {
								$email = $row[self::iEmail];
							}
						}
					//Check if this slip already assigned to this unit and merge last name
					$sql = "SELECT * FROM Slips WHERE slipid = ?" ;
					$stmt = $this->pdo->prepare ($sql);
					$stmt->execute([trim($slip)]);
					if ($stmt->fetchColumn()) {
						//found slip exists so check if same unit
						
					} else {
						//setup SQL statement for insert to Slips table
						$sql = "INSERT INTO Slips (unit, names, slipid, lift,phone,email)
								VALUES ('" . $row[7] . "', '" . $row[6] . "', '" . trim($slip) . "', '" . $lift . "', '" . $phone . "', '" . $email . "')";
						// execute the SQL statement - if returns fail then report
						if ($this->pdo->query($sql)){
						} else {
							echo "Failed " . $slip . "<br>";
						}
					}
				}
			}
		}
		//now query to dbSlip
		$query1 = $this->pdo->prepare("SELECT dock, b.slipid, b.class, rate, b.type, b.condition, a.names, unit, lift, phone, email
							FROM SlipMaster b
							LEFT OUTER JOIN Slips a ON a.slipid = b.slipid
							JOIN RateMaster c ON b.class = c.class
							WHERE b.type = 'Slip' ORDER BY b.slipid");
		$query1->execute();
		$this->dbSlip = $query1->fetchALL(PDO::FETCH_ASSOC);
				
		//now add waitlist info at bottom
		$this->dbSlip[]=array("");
		$this->dbSlip[]=array("Wait List");
		$this->dbSlip[]=array_keys($this->hdrWait);
		$query1 = $this->pdo->prepare("SELECT date, names, unit FROM WaitList where type = 'S' ORDER BY date");
		$query1->execute();
		$this->dbSlip = array_merge($this->dbSlip,$query1->fetchALL(PDO::FETCH_ASSOC));
		
		
		//now query to dbKayak
		$query1 = $this->pdo->prepare("SELECT dock, b.slipid, b.class, rate, b.type, b.condition, a.names, unit, lift, phone, email 
							FROM SlipMaster b
							LEFT OUTER JOIN Slips a ON a.slipid = b.slipid
							JOIN RateMaster c ON b.class = c.class
							WHERE b.type = 'Kayak' ORDER BY b.slipid");
		$query1->execute();
		$this->dbKayak = $query1->fetchALL(PDO::FETCH_ASSOC);
		$this->dbKayak[]=array("");
		$this->dbKayak[]=array("Wait List");
		$this->dbKayak[]=array_keys($this->hdrWait);
		$query1 = $this->pdo->prepare("SELECT date, names, unit, number FROM WaitList where type = 'K' ORDER BY date");
		$query1->execute();
		$this->dbKayak = array_merge($this->dbKayak,$query1->fetchALL(PDO::FETCH_ASSOC));
		
		return;
	}

	function BuildGrids()
	{
		//Build the resident grid
		//setup grid headers
		$this->hdrGridT1 = array("Floor"=>"string","Rembrandt-1"=>"string", "Monet-2"=>"string", "Renoir-3"=>"string", "Van Gogh-4"=>"string", "Van Gogh-5"=>"string", "Renoir-6"=>"string", "Monet-7"=>"string", "Cezanne-8"=>"string");

		$this->hdrGridT2 = array("Floor"=>"string","Rembrandt-9"=>"string", "Renoir-10"=>"string", "Renoir-11"=>"string", "Van Gogh-12"=>"string", "Van Gogh-14"=>"string", "Renoir-15"=>"string", "Renoir-16"=>"string", "Rembrandt-17"=>"string");

		//get owner, renter from db
		$j=0;
		$i=1;
		$query1 = $this->pdo->prepare("SELECT Owner, Renter, Floor, Stack
							FROM UnitMaster 
							WHERE Bldg LIKE 'Tower%'
							ORDER BY Floor DESC, Stack ASC");
		$query1->execute();
		// now loop thru all units
		while ($row = $query1->fetch(PDO::FETCH_ASSOC)) {
			$this->dbGridT1[$j][0] = $row['Floor'];
			$this->dbGridT2[$j][0] = $row['Floor'];
			if ($row['Stack']<=8) {
				$this->dbGridT1[$j][$row['Stack']] = $this->AddToGrid($row);
			} elseif ($row['Stack'] == 15 && $row['Floor'] == 1) {
				$this->dbGridT2[$j][6] = " ";
				$this->dbGridT2[$j][7] = $this->AddToGrid($row);
			} elseif ($row['Stack']>12) {
				$this->dbGridT2[$j][$row['Stack']-7] = $this->AddToGrid($row);
			} else {
				$this->dbGridT2[$j][$row['Stack']-8] = $this->AddToGrid($row);
			}
			$i++;
			if ($i>16) {
				$i=1;
				$j++;
			}
		}
	}

	function AddToGrid($row)
	{
		//add owners to front of string and renters in () to back of string and build vVoter array
		$r = $row['Owner'];
		if (strlen($row['Renter'])>0) {
			$r .= " (" . $row['Renter'] . ")";
		}
		return $r;
	}

	function CreateFile() 
	{
		// Include the required Class file
		include_once('PAI_xlsxwriter.class.php');
		$filename = "UserAddresses.xlsx";
		header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		
		//setup heading row style
		$hstyle = array( 'font'=>'Arial','font-size'=>10,'font-style'=>'bold', 'halign'=>'center', 'border'=>'bottom');
		$h1style = array( 'font'=>'Arial','font-size'=>12,'font-style'=>'bold', 'halign'=>'left', 'border'=>'bottom');
		
		//sort
		sort($this->dbErr);
		sort($this->dbPets);
		sort($this->dbRenter);

		//write header then sheet data and output file
		$writer = new XLSXWriter();
		$writer->setAuthor('Chris Barlow');
		if ($this->fullRun) {
			$writer->setColWidths('Errors',array(10,20,30,20,40));
			$writer->writeSheetHeader('Errors',$this->hdrErr,true);
			$writer->writeSheetRow('Errors',array(date('m/d/y'),'Condo on the Bay Error Listing'),$h1style);
			$writer->writeSheetRow('Errors',array_keys($this->hdrErr),$hstyle);
			$writer->writeSheet($this->dbErr,'Errors',$this->hdrErr,true);
			
			$writer->setColWidths('Listing',array(20,15,20,15,15,30,30,30));
			$writer->writeSheetHeader('Listing',$this->hdrRes,true);
			$writer->writeSheetRow('Listing',array(date('m/d/y'),'Condo on the Bay Owner & Renter Listing'),$h1style);
			$writer->writeSheetRow('Listing',array_keys($this->hdrRes),$hstyle);
			$writer->writeSheet($this->dbRes,'Listing',$this->hdrRes,true);
			
			$writer->setColWidths('Renter',array(12,12,20,15,15,15,15,30,30,30));
			$writer->writeSheetHeader('Renter',$this->hdrRenter,true);
			$writer->writeSheetRow('Renter',array(date('m/d/y'),'Condo on the Bay Renter Listing'),$h1style);
			$writer->writeSheetRow('Renter',array_keys($this->hdrRenter),$hstyle);
			$writer->writeSheet($this->dbRenter,'Renter',$this->hdrRenter,true);
			
			$writer->setColWidths('Pets WSD-ESA',array(20,15,15,50,15,15,30,30,30));
			$writer->writeSheetHeader('Pets WSD-ESA',$this->hdrPets,true);
			$writer->writeSheetRow('Pets WSD-ESA',array(date('m/d/y'),'Condo on the Bay Pets & WSD/ESA Listing'),$h1style);
			$writer->writeSheetRow('Pets WSD-ESA',array_keys($this->hdrPets),$hstyle);
			$writer->writeSheet($this->dbPets,'Pets WSD-ESA',$this->hdrPets,true);
			
			$writer->setColWidths('Slips',array(12,10,15,10,6,15,15,20,6,20,40));
			$writer->writeSheetHeader('Slips',$this->hdrSlip,true);
			$temp = ($this->showInfo) ? 'Internal':'External';
			$writer->writeSheetRow('Slips',array(date('m/d/y'),$temp,'Condo on the Bay Slip Listing'),$h1style);
			$writer->writeSheetRow('Slips',array_keys($this->hdrSlip),$hstyle);
			$writer->writeSheet($this->dbSlip,'Slips',$this->hdrSlip,true);
			
			$writer->setColWidths('Kayak',array(12,10,15,10,6,15,15,20,6,20,40));
			$writer->writeSheetHeader('Kayak',$this->hdrSlip,true);
			$writer->writeSheetRow('Kayak',array(date('m/d/y'),$temp,'Condo on the Bay Kayak Listing'),$h1style);
			$writer->writeSheetRow('Kayak',array_keys($this->hdrSlip),$hstyle);
			$writer->writeSheet($this->dbKayak,'Kayak',$this->hdrSlip,true);
			
			$writer->setColWidths('Staff',array(20,15,40,40));
			$writer->writeSheetHeader('Staff',$this->hdrStaff,true);
			$writer->writeSheetRow('Staff',array(date('m/d/y'),'Condo on the Bay Staff Listing'),$h1style);
			$writer->writeSheetRow('Staff',array_keys($this->hdrStaff),$hstyle);
			$writer->writeSheet($this->dbStaff,'Staff',$this->hdrStaff,true);
			
			$writer->setColWidths('Grid T1',array(10,20,20,20,20,20,20,20,20));
			$writer->writeSheetHeader('Grid T1',$this->hdrGridT1,true);
			$writer->writeSheetRow('Grid T1',array(date('m/d/y'),'Condo on the Bay T1 Grid'),$h1style);
			$writer->writeSheetRow('Grid T1',array_keys($this->hdrGridT1),$hstyle);
			$writer->writeSheet($this->dbGridT1,'Grid T1',$this->hdrGridT1,true);
			
			$writer->setColWidths('Grid T2',array(10,20,20,20,20,20,20,20,20));
			$writer->writeSheetHeader('Grid T2',$this->hdrGridT2,true);
			$writer->writeSheetRow('Grid T2',array(date('m/d/y'),'Condo on the Bay T2 Grid'),$h1style);
			$writer->writeSheetRow('Grid T2',array_keys($this->hdrGridT2),$hstyle);
			$writer->writeSheet($this->dbGridT2,'Grid T2',$this->hdrGridT2,true);
		}

		$writer->setColWidths('Voters',array(30,40,30,20,15));
		$writer->writeSheetHeader('Voters',$this->hdrVoter,true);
		$writer->writeSheetRow('Voters',array_keys($this->hdrVoter),$hstyle);
		$writer->writeSheet($this->dbVoter,'Voters',$this->hdrVoter,false);
		
		$writer->writeSheetHeader('Users',$this->hdrUser,true);
		$writer->writeSheetRow('Users',array_keys($this->hdrUser),$hstyle);
		$writer->writeSheet($this->dbUser,'Users',$this->hdrUser,false);
		
		if ($this->fullRun) {
			$writer->writeSheetRow('Sold',array_keys($this->hdrUser),$hstyle);
			$writer->writeSheet($this->dbSold,'Sold',$this->hdrUser,true);
			
			$writer->writeSheetRow('Gone',array_keys($this->hdrUser),$hstyle);
			$writer->writeSheet($this->dbGone,'Gone',$this->hdrUser,true);
		}
		$writer->writeToStdOut(); 
		unset($writer);
		return;	
	}

	

// ------ functions called by routines above -------------------------------
	
	function GetAddress($row) {
	// returns an array of addr, citystate for mailings based on user settings
	if ($row[self::iMailings] == "2nd Address") {
		if ($temp = explode (',',$row[self::i2ndAddress])) {
			if (count($temp)!==2) {
				$temp[0]  = "";
				$temp[1]  = "";
				$this->addError('4','2ndAddress format',$row[self::iUnit],$row[self::iUser1LastName],$row[self::i2ndAddress]);
			}
		}
	} else {
		switch (substr($row[self::iUnit],0, 7)) {
			Case "Tower 1":
				$temp[0] = "888 Blvd of the Arts " . substr($row[self::iUnit],8, 5);
				$temp[1] = "Sarasota FL 34236";
				break;
			Case "Tower 2":
				$temp[0] = "988 Blvd of the Arts " . substr($row[self::iUnit],8, 5);
				$temp[1] = "Sarasota FL 34236";
				break;
			Case "Marina ":
				$temp[0] = substr($row[self::iUnit],16, 3) . " Blvd of the Arts";
				$temp[1] = "Sarasota FL 34236";
				break;
			default:
				$this->addError('1','Unit format',$row[self::iUnit],$row[self::iUser1LastName],'Unit has wrong association');
				$temp[0] = "";
				$temp[1]  = "";
			}	
		}
	return $temp;
	}
	
	function GetVoter($row){
		// gets the name of unit voter
			If ($row[self::iOfficialVoter] == "Resident1") {
				$R = $row[self::iFirstName] . " " . $row[self::iUser1LastName];
			} elseif ($row[self::iOfficialVoter] == "Resident2") {
				$R = $row[self::iUser2FirstName] . " " . $row[self::iUser2LastName];
			} else {
				$R="";
			}
		return $R;
	}
	
	function GetFloor($unit)
	{
	// REPLACE WITH DB QUERY?
	// return the floor based in the unit. for multi return both floors
	// change to get floor from UnitMaster
	//receives unit string as return common delimited floor(s)
	//unit can be Tower 1 # 708 or Tower 1 #1706/1103 or Tower 1 #1903/04/05 or Tower 1 #1003/Tower 2 # 202

	$Units = explode('/', $unit,5);
	$F= "";
	foreach ($Units as $unit) {
		$unit = trim($unit);
		switch (strlen($unit)) {
			Case 19: //MS
				if (strlen($F)>0) {
					$F = $F . ",1";
				} else {
					$F = "1";
				}
				break;
			Case 13: //full unit
				if (strlen($F)>0) {
					$F = $F . "," . trim(substr($unit, 9, 2));
				} else {
					$F = trim(substr($unit, 9, 2));
				}
				break;
			Case 14:
				if (strlen($F)>0) {
					$F = $F . "," . substr($unit, 10, 2);
				} else {
					$F = substr($unit, 10, 2);
				}
				break;
			Case 4: //unit only
				if (strlen($F)>0) {
					$F = $F . "," . substr($unit, 0, 2);
				} else {
					$F = substr($unit, 1, 2);
				}
				break;
			Case 3: //unit only
				if (strlen($F)>0) {
					$F = $F . "," . substr($unit, 0, 1);
				} else {
					$F = substr($unit, 1, 1);
				}
				break;
			Case 2: //same floor as primary
				$F = $F . "," . substr($F, 0, 2);
				break;
			Case 1:
				$F = $F . "," . substr($F, 0, 1);
				break;
			default:
				$F = strlen($unit);
				$this->addError('1','Unit format',$unit,'','Incorrect format - could not calc floor');
			}
		}
		return $F;
	}
	
	function LogRun()
	{
		//update RunLog table for this run
		$ip = $_SERVER['REMOTE_ADDR'] ;
		$sql = "INSERT INTO RunLog (ip,type,records, showinfo)
				VALUES ('" . $ip  . "','1','" . count($this->dbUser) . "', '" . boolval($this->showInfo) . "')";
		// execute the SQL statement - if returns fail then report
		$this->pdo->query($sql);
		
		//now email
		$to      = 'cbarlow@pathfinderassociatesinc.com';
		$subject = 'COBList run';
		$message = $sql;
		$headers = 'From: webmaster@condoonthebay.com' . "\r\n" .
			'Reply-To: webmaster@condoonthebay.com' . "\r\n" .
			'X-Mailer: PHP/' . phpversion();
		mail($to, $subject, $message, $headers);
		
		error_log($sql,1,$to);
		
		return ;
	}

	function addError($level, $function, $unit, $name, $message)
	{
		if ((stripos($unit,'gone') !== false) || (stripos($unit,'sold') !== false)) {
		} else {
			$tmp = array(
				'level' => $level,
				'function'	=> $function,
				'unit'	=> $unit,
				'name'	=> $name,
				'message'	=> $message,
				);
			$this->dbErr[] = $tmp;
		}
	return;
	}

function opendb(&$checkmsg) {
	//function to open PDO database and return PDO object
	$host = 'localhost';
	$db   = 'coblist';
	$user = 'cobuser';
	$pass = 'sarasota888';
	$charset = 'utf8';

	$dsn = "mysql:host=$host;dbname=$db;charset=$charset";
	$opt = [
		PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
		PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
		PDO::ATTR_EMULATE_PREPARES   => false,
	];
	
	try {
		$this->pdo = new PDO($dsn, $user, $pass, $opt);
	} catch (PDOException $e) {
		$checkmsg = 'Connection failed: ' . $e->getMessage();
		return false;
	}
	return true;
}

function CheckUnitFormat ($row) {
	// check unit format against unit master and update owner, owner count, renter, voter fields
	if (strlen($row[self::iUnit])>0) {
		// explode unit
//		$temp = $this->GetFullUnit($row[self::iUnit]);

		if ($temp = $this->GetFullUnit($row[self::iUnit])) {
		$S="";
		foreach ($temp as $u) {
			//check each unit in UnitMaster and confirm Assoc, Floor, Stack
			$sql = "SELECT * FROM UnitMaster WHERE Unit = ?" ;
			$stmt = $this->pdo->prepare ($sql);
			$stmt->execute([trim($u)]);
			$result = $stmt->fetch();
			if ($result) {
				if (stripos($row[self::iOwner],'yes')!== false) { 
					//if different owner then increment count in UnitMaster
					if (stripos($result['Owner'],$row[self::iUser1LastName]) !== false) {
					} else {
						$sql = "UPDATE UnitMaster SET Owner = ?, count = count + 1 WHERE Unit = ?" ;
						$stmt = $this->pdo->prepare ($sql);
						if (strlen($result['Owner'])>0) {
							$temp = $result['Owner'] . '/' . $row[self::iUser1LastName];
						} else {
							$temp = $row[self::iUser1LastName];
							}
						$stmt->execute([$temp,trim($u)]);
						}
					//if different voter then update in UnitMaster
					if (stripos($result['Voter'],$row[self::iVoter]) !== false) {
					} elseif (strlen(trim($row[self::iVoter]))>0) {
						$sql = "UPDATE UnitMaster SET Voter = ?, Address = ?, CityStateZip = ? WHERE Unit = ?" ;
						$stmt = $this->pdo->prepare ($sql);
						if (strlen(trim($result['Voter']))>0) {
							$temp = $result['Voter'] . '/' . $row[self::iVoter];
						} else {
							$temp = $row[self::iVoter];
							}
						$stmt->execute([$temp,$row[self::iAddress],$row[self::iCityStateZip],trim($u)]);
						}
					switch ($result['Bldg']) {
					Case "Tower 1":
						if (strpos($row[self::iAccess], '^T1') === false) {
							$this->addError('1','Owner access',$u,$row[self::iUser1LastName],'T1 Owner without ^T1 access');
						}
						//calculate stack
						$S .= $result['Stack'] . ", ";
						break;
					Case "Tower 2":
						if (strpos($row[self::iAccess], '^T2') === false) {
							$this->addError('1','Owner access',$u,$row[self::iUser1LastName],'T2 Owner without ^T2 access');
						}
						$S .= $result['Stack'] . ", ";
						break;
					Case "Marina ":
						if (strpos($row[self::iAccess], '^MS') === false) {
							$this->addError('1','Owner access',$u,$row[self::iUser1LastName],'MS Owner without ^MS access');
						}
						$S .= $result['Stack'] . ", ";
						break;
					}
				} elseif (stripos($row[self::iOwner],"no")!== false) {
					//if different renter then update in UnitMaster
					if (stripos($result['Renter'],$row[self::iUser1LastName]) !== false) {
					} else {
						$sql = "UPDATE UnitMaster SET Renter = ? WHERE Unit = ?" ;
						$stmt = $this->pdo->prepare ($sql);
						if (strlen($result['Renter'])>0) {
							$temp = $result['Renter'] . '/' . $row[self::iUser1LastName];
						} else {
							$temp = $row[self::iUser1LastName];
							}
						$stmt->execute([$temp, trim($u)]);
					}
					
					//renters should not have owner access
					if (!($row[self::iAccess] == "MEMBER") ) {
						$this->addError('2','Renter error',$u,$row[self::iUser1LastName],'Renter with wrong access');
					}
					$S .= $result['Stack'] . ", ";
				} else {
					//Owner not Yes or No
						$this->addError('2','Owner field',$u,$row[self::iUser1LastName],'Owner field not Yes/No');
				}
			} else {
				// unit not in UnitMaster
					$this->addError('2','Unit error',$u,$row[self::iUser1LastName],'Unit not valid format');
				
			}
		}
		//check stack after removing last ,
		$S = substr($S, 0, -2);
		if ($row[self::iStack] !== $S && $row[self::iAccess] !== "PUBLIC"  && strlen($S)) { 
		//check if moved or sold by Access
			$temp = $row[self::iStack] . " should be " . $S;
			$this->addError('3','Stack error',$row[self::iUnit],$row[self::iUser1LastName],$temp);
		}
	}}
return;	
}

function CheckPhoneFormat($p) {
	if (strlen(trim($p))>0) {
		if (!preg_match('/^(\+1|001)?\(?([0-9]{3})\)?([ .-]?)([0-9]{3})([ .-]?)([0-9]{4})/',$p)) { 
//		if (!preg_match('\(?[2-9][0-8][0-9]\)?[-. ]?[1-9][0-9]{2}[-. ]?[0-9]{4}',$p)) {
			return false;
		} else {
			return true;
		}
	} else {
		return true;
	}
}

function GetBestPhone($row){
	// function returns cell or other phone
	if (strlen($row[10])>0) {
		return $row[10];
	} elseif (strlen($row[8])>0) {
		return $row[8];
	} elseif (strlen($row[9])>0) {
		return $row[9];
	} else {
		$this->addError('4','No phone number',$row[self::iUnit],$row[self::iUser1LastName],'No phone number for this resident');
		return "";
	}
}

function GetLeaseDates($row) {
	// return array of start date and end date and log error
	$D = explode('-', $row[self::iOfficialVoter]);
	if (count($D) !== 2) {
		$temp = 'Wrong lease date format=' . $row[self::iOfficialVoter];
		$this->addError('5','Lease dates',$row[self::iUnit],$row[self::iUser1LastName],$temp);
		$D[1] = "Error";
	} else {
		$D[0] = date_format(date_create_from_format("m/d/y",$D[0]),"Y-m-d");
		$D[1] = date_format(date_create_from_format("m/d/y",$D[1]),"Y-m-d");
	}
	return $D;
}

function GetFullUnit($unit) {
	// this returns an array of full units from abbreviated units
	$Units = array_filter(explode('/', $unit));
	$F = "";
	$T = "";
	foreach ($Units as $u) {
		$u = trim($u);
		switch (strlen($u)) {
			Case 19: //MS Marina Suites # 902
				$F[]=$u;
				$T = substr($u,0,15);
				break;
			Case 13: //full u Tower 1 # 708
				$F[]=$u;
				$T = substr($u,0,11);
				break;
			Case 4: //u only 1801
				$F[]= substr($T,0,9) . $u;
				break;
			Case 3: //u only 708
				$F[]= substr($T,0,9) . " " . $u;
				break;
			Case 2: //same floor as primary 02 or 12
				$F[]= $T . $u;
				break;
			Case 1: // 3
				$F[]= $T . "0" . $u;
				break;
			default:
				$this->addError('1','Unit format',$unit,count($F),'Incorrect format - could not calc floor');
				return false;
			}
		}
	return $F;
}
// end of class	
}
?>