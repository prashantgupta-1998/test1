package com.sampleproject.sample_project.Interface;

import java.io.ByteArrayInputStream;

public interface EmployeeInterface {

    default ByteArrayInputStream downloadAllUserData() {
        return null;
    }

}
