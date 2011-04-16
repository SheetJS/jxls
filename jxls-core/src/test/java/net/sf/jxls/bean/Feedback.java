package net.sf.jxls.bean;

import java.io.Serializable;
import java.util.Calendar;
import java.util.Date;

/**
 * @author Graham Rhodes 3 Mar 2011 12:27:41
 */
public class Feedback implements Serializable, Comparable<Feedback> {

    private static final long serialVersionUID = 1L;
    public static final String RECEIVER = "receiver";
    public static final String CREATED = "created";
    public static final String RATING = "rating";
    private int feedId = 0;
    private String receiver = null;
    private int rating = 0;
    private String comment = "";
    private String createdBy = null;
    private Date created = new Date(Calendar.getInstance().getTimeInMillis());
    private String updatedBy = null;
    private Date lastUpdated = new Date(Calendar.getInstance().getTimeInMillis());

    public Feedback() {

    }

    public Feedback(int rating, Date created, String comment, String creator, String receiver) {
        this.created = created;
        this.rating = rating;
        this.comment = comment;
        this.receiver = receiver;
        this.createdBy = creator;
        this.updatedBy = creator;
    }

    public int getFeedId() {
        return feedId;
    }

    public String getReceiver() {
        return receiver;
    }

    public int getRating() {
        return rating;
    }

    public String getComment() {
        return comment;
    }

    public String getCreatedBy() {
        return createdBy;
    }

    public Date getCreated() {
        return created;
    }

    public String getUpdatedBy() {
        return updatedBy;
    }

    public Date getLastUpdated() {
        return lastUpdated;
    }

    public void setFeedId(int feedId) {
        this.feedId = feedId;
    }

    public void setReceiver(String receiver) {
        this.receiver = receiver;
    }

    public void setRating(int rating) {
        this.rating = rating;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    public void setCreatedBy(String createdBy) {
        this.createdBy = createdBy;
    }

    public void setCreated(Date created) {
        this.created = created;
    }

    public void setUpdatedBy(String updatedBy) {
        this.updatedBy = updatedBy;
    }

    public void setLastUpdated(Date lastUpdated) {
        this.lastUpdated = lastUpdated;
    }

    public int compareTo(Feedback other) {
        return created.compareTo(other.getCreated());
    }

}
