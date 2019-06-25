grid = null;
features2sync = {};

function strip(html)
{
   var tmp = document.createElement("DIV");
   tmp.innerHTML = html;
   return tmp.textContent || tmp.innerText || "";
}

function getIteration(iterationPath){
	return iterationPath.substring(iterationPath.lastIndexOf("\\")+1);
}

function getQuarter(iterationPath){
	return iterationPath.substring(0, iterationPath.lastIndexOf("\\"));
}

function loadChanges(workItemsResults){
	for (var i=0; i < workItemsResults.length; i++){	
		let r = workItemsResults[i];	
		features2sync[r.id] = {Id:r.id, Fields:r.fields, Comment:null,Links:[],Parent:null}
		
		var commentsLoaded = 0;
		TFSWitWebApi.getClient().getComments(r.id)
		.then(function (comments) {
			for (var k=0; k < comments.count; k++){
				var commentHtmlText = comments.comments[k].text;
				if(commentHtmlText){
					var commentText = strip(commentHtmlText);
					var commentChangeIndexOf = commentText.indexOf('CHANGE:');
					if(commentChangeIndexOf!=-1){
						var commentJsonText = commentText.substring(commentChangeIndexOf+7);
						var commentJson	= JSON.parse(commentJsonText);
						var commnetJsonFrom = commentJson.From;
						commentJson.From = getIteration(commnetJsonFrom);
						var commnetJsonTo = commentJson.To;
						commentJson.To = getIteration(commnetJsonTo);

						features2sync[r.id].Comment = commentJson;
					}
				}
			}

			if (++commentsLoaded == workItemsResults.length){
				loadParents(workItemsResults);
			}
		});
	}
}

function loadParents(workItemsResults){
	var parents = [];
	for (var i=0; i < workItemsResults.length; i++){
		let r = workItemsResults[i];
		for (var j=0; j < r.relations.length; j++){
			var relation = r.relations[j];
			if(relation.rel === "System.LinkTypes.Hierarchy-Reverse"){				
				var parentUrl = relation.url;
				var parentId = parentUrl.substring(parentUrl.lastIndexOf("/")+1);
				features2sync[r.id].Parent = {Id: parentId,Name:"",Type:""}
				parents.push(parentId);
			}
		}
	}
	
	if(parents.length != 0){
		TFSWitWebApi.getClient().getWorkItems(parents)
		.then(function (parentsResults) {
			for (var k=0; k < parentsResults.length; k++){
				var parentsResult = parentsResults[k];				
				for (var key in features2sync) {
					var parent = features2sync[key].Parent;
					if (parent.Id == parentsResult.id){
						parent.Name = parentsResult.fields["System.Title"];
						parent.Type = parentsResult.fields["System.WorkItemType"];
					}
				}
			}
			
			loadDependencies(workItemsResults);
		});
	}else{
		loadDependencies(workItemsResults);
	}
}

function loadDependencies(workItemsResults){
	var dependencies = [];
	for (var i=0; i < workItemsResults.length; i++){
		let r = workItemsResults[i];
		for (var j=0; j < r.relations.length; j++){
			var relation = r.relations[j];
			if(relation.rel === "System.LinkTypes.Dependency-Reverse"){				
				var dependencyUrl = relation.url;
				var dependencyId = dependencyUrl.substring(dependencyUrl.lastIndexOf("/")+1);
				dependencies.push(dependencyId);
				
				var link = {Id: r.id,Severity:"", DueDate:"", Team:"", DependencyId:dependencyId,r:null};
				var commentRel = relation.attributes.comment;
				if (commentRel){
					var comment = commentRel.split(":");
					if (comment.length == 2){
						link.Severity = comment[0];
						link.DueDate = getIteration(comment[1]).replace("'","");
					}
				}
				
				features2sync[r.id].Links.push(link);
			}
		}
	}
	
	if(dependencies.length != 0){
		TFSWitWebApi.getClient().getWorkItems(dependencies)
		.then(function (dependenciesResults) {
			for (var k=0; k < dependenciesResults.length; k++){
				var dependencyResult = dependenciesResults[k];
				
				for (var key in features2sync) {
					var links = features2sync[key].Links;
					if (links){
						for (var x=0; x < links.length; x++){
							var link = links[x];
							if (link.DependencyId == dependencyResult.id){
								link.Fields = dependencyResult.fields;
							}
						}
					}
				}
			}
			
			getColors();
		});
	}else{
		getColors();
	}
}

workItemTypesStyles = {};
function getColors(){
	TFSWitWebApi.getClient().getWorkItemTypes(VSS.getWebContext().project.name)
	.then(function (workItemIcons) {
		for (var i=0; i < workItemIcons.length; i++){
			var item = workItemIcons[i];
			statesStyles = {};
			for (var j=0; j < item.states.length; j++){
				var state = item.states[j];
				statesStyles[state.name] = state.color;
			}

			workItemTypesStyles[item.name] = {color:item.color, iconUrl:item.icon.url, statesStyles:statesStyles};
		}
		
		convertToGridSource();
	});
}

function convertToGridSource(){
	var gridSource = [];
	for (var key in features2sync) {
		var feature = features2sync[key];
		var reason = null;
		if (feature.Comment){
			reason = feature.Comment.Reason;
		}

		var gridSourceItem = {
			pid:feature.Parent.Id,
			pt:feature.Parent.Type,
			pn:feature.Parent.Name,
			fid:feature.Id,
			fn:feature.Fields["System.Title"],
			ft:feature.Fields["System.WorkItemType"],
			fs:feature.Fields["System.State"],
			fr:reason,
			fpi:getIteration(feature.Fields["System.IterationPath"]),
			fu:(feature.Fields["System.Tags"] || "").includes("Unplanned"),
			lt:"",
			ls:"",
			ld:"",
			children:[]
		};
		
		for (var i=0; i < feature.Links.length; i++){
			var link = feature.Links[i];
			var child = {	
				pid:feature.Id,			
				pt:null,
				pn:null,
				fid:link.Id,
				fn:link.Fields["System.Title"],
				ft:link.Fields["System.WorkItemType"],
				fs:link.Fields["System.State"],
				fr:null,
				fpi:getIteration(link.Fields["System.IterationPath"]),
				fu:false,
				lt:link.Fields["Payvision.Team"],
				ls:link.Severity,
				ld:link.DueDate
			};
			gridSourceItem.children.push(child);
		}
		
		gridSource.push(gridSourceItem);
	}
	
	drawTable(gridSource);
}

function drawTable(gridSource){	
	var gridOptions = {
		height: "100%",
		width: "100%",
		source: new VSSControlsGrids.GridHierarchySource(gridSource),
		sortOrder: [{ index: "fpi", order: "asc" }],
		columns: [
			{ text: "Planned", index: "fpi", width: 140, indent: true },
			{ text: "Delivered", index: "fdi", width: 120 , getCellContents: function (rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
					var gridCell = $("<div class='grid-cell'/>").width(column.width);
					gridCell.css("text-indent", (16 * (level-1)) + "px");

					var titleText = $("<div style='display:inline' />").text(this.getColumnValue(dataIndex, "fpi"));					
					gridCell.append(titleText);
					
					if (this.getColumnValue(dataIndex, "pn")){
						var finishDate = teamIterations[this.getColumnValue(dataIndex, "fpi")].startDate;
						var today = new Date();
						var state = this.getColumnValue(dataIndex, "fs");
						if (today > finishDate && (state === "Active" || state === "New")){
							rowInfo.row.context.style.backgroundColor = "#FF000022";
						}
						
						if (this.getColumnValue(dataIndex, "fu")){
							rowInfo.row.context.style.backgroundColor = "#0000FF22";							
						}
					}

					return gridCell;
				}
			},
			{ text: "State", index: "fs", width: 90, getCellContents: function (rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
					var indent = (16 * (level-1)) + "px";
					var gridCell = $("<div class='grid-cell' style='margin-left:"+indent+"'/>").width(column.width);
					var decorator = $("<div class='state-circle' style='float:left; border-radius:50%; width:10px; height: 10px; margin-top: 4px; margin-right: 5px;' />");
					
					var workItemType = this.getColumnValue(dataIndex, "ft");
					var state = this.getColumnValue(dataIndex, "fs");
					decorator.css("background-color", workItemTypesStyles[workItemType].statesStyles[state]);
					gridCell.append(decorator);
					
					var titleText = $("<div style='display:inline' />").text(state);					
					gridCell.append(titleText);

					return gridCell;
				}
			},
			//{ text: "Parent Type", index: "pt", width: 100 },
			{ text: "Epic", index: "pn", width: 350, getCellContents: function (rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
					var gridCell = $("<div class='grid-cell'/>").width(column.width);
					var workItemType = this.getColumnValue(dataIndex, "pt");
					if (!workItemType){
						return gridCell;						
					}
					
					var decorator = $("<img src='"+workItemTypesStyles[workItemType].iconUrl+"' width='14px' />");
					
					var titleHref = $("<a>");
					titleHref.on("click", () => {
						TFSWitServices.WorkItemFormNavigationService.getService().then(service => {
							service.openWorkItem(this.getColumnValue(dataIndex, "pid"), false);
						});
					});
					titleHref.text(this.getColumnValue(dataIndex, "pn"));
					var titleText = $("<div style='display:inline' />").add(titleHref);
					
					gridCell.append(decorator);
					gridCell.append(titleText);

					return gridCell;
				}
			},
			//{ text: "Type", index: "ft", width: 80 },
			{ text: "Title", index: "fn", width: 500, getCellContents: function (rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
					var gridCell = $("<div class='grid-cell'/>").width(column.width);
					gridCell.css("text-indent", (16 * (level-1)) + "px");
					var workItemType = this.getColumnValue(dataIndex, "ft");
					var decorator = $("<img src='"+workItemTypesStyles[workItemType].iconUrl+"' width='14px' />");
										
					var titleHref = $("<a>");
					titleHref.on("click", () => {
						TFSWitServices.WorkItemFormNavigationService.getService().then(service => {
							service.openWorkItem(this.getColumnValue(dataIndex, "fid"), false);
						});
					});
					titleHref.text(this.getColumnValue(dataIndex, "fn"));
					var titleText = $("<div style='display:inline' />").add(titleHref);
					
					gridCell.append(decorator);
					gridCell.append(titleText);

					return gridCell;
				}
			},
			{ text: "Dependency", index: "lt", width: 80 },
			{ text: "Others", index: "o", width: 300, getCellContents: function (rowInfo, dataIndex, expandedState, level, column, indentIndex, columnOrder) {
					var divContent = "";
					if(this.getColumnValue(dataIndex, "ls")){
						divContent = "<span style='font-weight:bold'>Severity: </span>" + this.getColumnValue(dataIndex, "ls") + " <span style='font-weight:bold'>from</span> " + this.getColumnValue(dataIndex, "ld");
					}else{
						var unplanned = null;
						if (this.getColumnValue(dataIndex, "fu")){
							unplanned = "Unplanned";
						}
						
						var reason = this.getColumnValue(dataIndex, "fr");
						if (unplanned || reason){
							divContent = "<span style='font-weight:bold'>Reason: </span>" + (unplanned || "") + (reason || "");
						}
					}
			
					return $("<div class='grid-cell'>" + divContent + "</div>").width(column.width || 300);
				}
			}
		]
	};
	
	grid = VSSControls.create(VSSControlsGrids.Grid, $("#mainContainer"), gridOptions);
}

teamIterations = {};
function CreateProgressChart(WidgetHelpers, TFS_Wit_WebApi, TFS_Wit_Contracts, TFS_Wit_Services, TFS_Work_WebApi, VSS_Controls, VSS_Controls_Grids) {	
	$("#mainContainer").empty();
	TFSWorkWebApi = TFS_Work_WebApi;
	TFSWitWebApi = TFS_Wit_WebApi;
	TFSWitServices = TFS_Wit_Services;
	VSSControls = VSS_Controls;
	VSSControlsGrids = VSS_Controls_Grids;
	teamContext = { projectId: VSS.getWebContext().project.id, teamId: VSS.getWebContext().team.id, project: "", team: "" }; 

		
	TFS_Work_WebApi.getClient().getTeamIterations(teamContext, "current")
	.then(function (iterations) {
		currentIteration = iterations[0];
		startDate = new Date(currentIteration.attributes.startDate.toDateString());
		finishDate = new Date(currentIteration.attributes.finishDate.toDateString());
		quarter = getQuarter(currentIteration.path);
		
		TFS_Work_WebApi.getClient().getTeamIterations(teamContext)
		.then(function (teamIterationsResult) {
			for (var i=0; i < teamIterationsResult.length; i++){
				var teamIteration = teamIterationsResult[i];
				var iterationQuarter = getQuarter(teamIterationsResult[i].path);
				if (iterationQuarter == quarter){
					teamIterations[teamIteration.name] = {startDate:teamIteration.attributes.startDate, finishDate:teamIteration.attributes.finishDate};
				}
			}
			
			var query = "select [System.Id]"
			query += " from WorkItems "
			query += " where [System.TeamProject] = 'B-Ops'"
			query += "  and [System.WorkItemType] = 'Feature'"
			query += "  and [System.State] <> 'Removed'";
			query += "  and [System.AreaPath] Under 'B-Ops' "
			query += "  and [System.IterationPath] Under '" + quarter + "'"
			query += " order by [System.Id]";

			var wiql = {query : query};	
			TFS_Wit_WebApi.getClient().queryByWiql(wiql, teamContext.projectId, teamContext.teamId)
			.then(function (queryResult) {			
				var allWorkItemIds = [];
				queryResult.workItems.forEach(element => {
					allWorkItemIds.push(element.id);
				});

				TFS_Wit_WebApi.getClient().getWorkItems(allWorkItemIds,null, null, TFS_Wit_Contracts.WorkItemExpand.Relations)
				.then(function (workItemsResults) {
					loadChanges(workItemsResults);
				});
			});	
		});
	});
}

VSS.require([
	"TFS/Dashboards/WidgetHelpers",
	"TFS/WorkItemTracking/RestClient",
	"TFS/WorkItemTracking/Contracts",
	"TFS/WorkItemTracking/Services",
	"TFS/Work/RestClient",
	"VSS/Controls",
	"VSS/Controls/Grids"
	], 
	
	CreateProgressChart
);