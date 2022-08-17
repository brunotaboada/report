package com.localhost.report;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class Record {
    private String name;
    private String ip;
    private String user;
    private String password;
    private String ipmiIp;
    private String ipmiUser;
    private String ipmiPassword;
}
