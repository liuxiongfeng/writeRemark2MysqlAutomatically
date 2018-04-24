/**
 * Created by yilun on 2017/5/23.
 */
public class Message {
    private boolean success;
    private String messageContent;

    public Message(boolean success) {
        this.success = success;
    }

    public Message() {}

    public String getMessageContent() {
        return messageContent;
    }

    public void setMessageContent(String messageContent) {
        this.messageContent = messageContent;
    }

    public boolean isSuccess() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }
}
