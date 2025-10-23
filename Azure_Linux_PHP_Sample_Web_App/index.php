 <!Doctype html>
<html>

<head>
     <title>AJ Labs PHP Application</title>
	 <link href="css/site.css" rel="stylesheet">
</head>

<body>
    <div class="main-container">
        <div class="cloud-image">
            <img src="img/successCloudNew.svg" />
        </div>
        <div class="content">
            <div class="tweet-container">
            <a href="https://www.linkedin.com/in/abdul-jaseem-ab1560114/">
                <img src="img/tweetThis.svg" />
            </a>            
        </div>
            <div class="content-body">
                <div class="success-text">Success!</div>
                <div class="description line-1"> Your Website is ready, HTML/PHP app is up and running!</div>
                
                <div class="next-steps-container">
                    <div class="next-steps-header">Next up</div>
                    <div class="next-steps-body">
                        <div class="step">
                            <div class="step-icon">
                                <img src="img/cloneWhite.svg">
                            </div>
                            <div class="step-text">Practice the task and get familiar with Web App Deployments</div>
                        </div>
                        <div class="step">
                            <div class="step-icon">
                                <img src="img/deployWhite.svg">
                            </div>
                            <div class="step-text"><a href="https://www.youtube.com/channel/UCgyu6yjUK_RCw6MgphtT-ig">More tutorials and videos on my YouTube channel</a></div>
                        </div>
                        <div class="step">
                            <div class="step-icon">
                                <img src="img/stackWhite.svg">
                            </div>
                            <div class="step-text"><a href="https://aj-labs.com/" target="_blank">Connect me</a> </div>
                        </div>
                        <div class="step">
                            <div class="step-icon">
                                <img src="img/lightbulbWhite.svg">
                            </div>
                            <div class="step-text">This website runs on Local IP:
								<span class="dynamic"> <?php $ip = $_SERVER['SERVER_ADDR']; echo "$ip"?> </span>
								and Host Name:
								<span class="dynamic"> <?php $hostname = exec("hostname");echo "$hostname"?> 
								</span>
							</div>
						</div>

						<div class="step">
                            <div class="step-icon">
                                <img src="img/lightbulbWhite.svg">
                            </div>
                            <div class="step-text">Other AJ Labs Trainings
								
								<ul class="step-text">
									<li><a href="https://aj-labs.com/courses/advanced-genesys-cloud-contact-center/Advanced_Genesys_Cloud_Contact_Center_by_AJ_Labs_Syllabus.pdf" target="_blank" rel="noopener">Adv. Genesys Cloud CX</a></li>
									<li><a href="https://aj-labs.com/courses/advanced-microsoft-teams-copilot-graph-api/Advanced_Microsoft_Teams_SBCs_Copilot_Graph_API_by_AJ_Labs_Syllabus.pdf" target="_blank" rel="noopener">Adv. Microsoft Teams</a></li>
									<li><a href="https://www.udemy.com/course/microsoft-teams-direct-routing-with-all-sbcs-part-of-ms-721/?referralCode=43FA284904FB9F8FC80D" target="_blank" rel="noopener">Microsoft Teams Direct Routing Fast Track</a></li>
									<li><a href="https://aj-labs.com/courses/advanced-cisco-collaboration/Advanced%20Cisco%20Collaboration%20Syllabus.pdf" target="_blank" rel="noopener">Adv. Cisco Collaboration</a></li>
								</ul>
							</div>
						</div>						
						
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
