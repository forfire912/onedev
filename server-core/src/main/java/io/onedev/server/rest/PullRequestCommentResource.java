package io.onedev.server.rest;

import io.onedev.server.entitymanager.PullRequestCommentManager;
import io.onedev.server.model.PullRequestComment;
import io.onedev.server.rest.annotation.Api;
import io.onedev.server.security.SecurityUtils;
import org.apache.shiro.authz.UnauthorizedException;

import javax.inject.Inject;
import javax.inject.Singleton;
import javax.validation.constraints.NotNull;
import javax.ws.rs.*;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import java.util.ArrayList;

@Api(order=3200)
@Path("/pull-request-comments")
@Consumes(MediaType.APPLICATION_JSON)
@Produces(MediaType.APPLICATION_JSON)
@Singleton
public class PullRequestCommentResource {

	private final PullRequestCommentManager commentManager;

	@Inject
	public PullRequestCommentResource(PullRequestCommentManager commentManager) {
		this.commentManager = commentManager;
	}

	@Api(order=100)
	@Path("/{commentId}")
	@GET
	public PullRequestComment get(@PathParam("commentId") Long commentId) {
		PullRequestComment comment = commentManager.load(commentId);
    	if (!SecurityUtils.canReadCode(comment.getProject()))  
			throw new UnauthorizedException();
    	return comment;
	}
	
	@Api(order=200, description="Update pull request comment of specified id in request body, or create new if id property not provided")
	@POST
	public Long createOrUpdate(@NotNull PullRequestComment comment) {
    	if (!SecurityUtils.canReadCode(comment.getProject()) || 
    			!SecurityUtils.isAdministrator() && !comment.getUser().equals(SecurityUtils.getUser())) { 
			throw new UnauthorizedException();
    	}

		if (comment.isNew())
			commentManager.create(comment, new ArrayList<>());
		else
			commentManager.update(comment);
		
		return comment.getId();
	}
	
	@Api(order=300)
	@Path("/{commentId}")
	@DELETE
	public Response delete(@PathParam("commentId") Long commentId) {
		PullRequestComment comment = commentManager.load(commentId);
    	if (!SecurityUtils.canModifyOrDelete(comment)) 
			throw new UnauthorizedException();
		commentManager.delete(comment);
		return Response.ok().build();
	}
	
}
