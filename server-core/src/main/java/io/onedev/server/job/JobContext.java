package io.onedev.server.job;

import io.onedev.k8shelper.Action;
import io.onedev.k8shelper.LeafFacade;
import io.onedev.k8shelper.ServiceFacade;
import io.onedev.server.OneDev;
import io.onedev.server.entitymanager.ProjectManager;
import io.onedev.server.model.Project;
import io.onedev.server.model.support.administration.jobexecutor.JobExecutor;
import org.eclipse.jgit.lib.ObjectId;

import java.io.Serializable;
import java.util.List;

public class JobContext implements Serializable {
	
	private static final long serialVersionUID = 1L;

	private final String jobToken;
	
	private final JobExecutor jobExecutor;

	private final Long projectId;
	
	private final String projectPath;
	
	private final String projectGitDir;
	
	private final Long buildId;
	
	private final Long buildNumber;

	private final Long submitSequence;
	
	private final List<Action> actions;
	
	private final String refName;
	
	private final ObjectId commitId;
	
	private final List<ServiceFacade> services;
	
	private final long timeout;

	public JobContext(String jobToken, JobExecutor jobExecutor, Long projectId, String projectPath,
					  String projectGitDir, Long buildId, Long buildNumber, Long submitSequence,
					  List<Action> actions, String refName, ObjectId commitId, List<ServiceFacade> services,
					  long timeout) {
		this.jobToken = jobToken;
		this.jobExecutor = jobExecutor;
		this.projectId = projectId;
		this.projectPath = projectPath;
		this.projectGitDir = projectGitDir;
		this.buildId = buildId;
		this.buildNumber = buildNumber;
		this.submitSequence = submitSequence;
		this.actions = actions;
		this.refName = refName;
		this.commitId = commitId;
		this.services = services;
		this.timeout = timeout;
	}
	
	public String getJobToken() {
		return jobToken;
	}

	public JobExecutor getJobExecutor() {
		return jobExecutor;
	}

	public Long getBuildId() {
		return buildId;
	}

	public List<Action> getActions() {
		return actions;
	}

	public String getRefName() {
		return refName;
	}

	public ObjectId getCommitId() {
		return commitId;
	}
	
	public List<ServiceFacade> getServices() {
		return services;
	}

	public Long getProjectId() {
		return projectId;
	}

	public String getProjectPath() {
		return projectPath;
	}

	public String getProjectGitDir() {
		return projectGitDir;
	}

	public Long getBuildNumber() {
		return buildNumber;
	}

	public Long getSubmitSequence() {
		return submitSequence;
	}

	public long getTimeout() {
		return timeout;
	}

	public LeafFacade getStep(List<Integer> stepPosition) {
		return LeafFacade.of(actions, stepPosition);
	}
	
	public boolean canManageProject(Project targetProject) {
		var project = OneDev.getInstance(ProjectManager.class).load(projectId);
		return project.isCommitOnBranch(commitId, project.getDefaultBranch())
				&& project.isSelfOrAncestorOf(targetProject);
	}
	
}
