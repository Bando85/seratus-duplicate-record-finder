/*
 * Copyright (c) 2023. 01. 04. 17:57. Created by Andras Laczo. All rights reserved.
 */

package andras.laczo.services;

public final class ExcelColumn {

        private ExcelColumn() {}

        public static int toNumber(String name) {
            int number = 0;
            for (int i = 0; i < name.length(); i++) {
                number = number * 26 + (name.charAt(i) - ('A' - 1));
            }
            return number;
        }

        public static String toName(int number) {
            number = number + 1;
            StringBuilder sb = new StringBuilder();
            while (number-- > 0) {
                sb.append((char)('A' + (number % 26)));
                number /= 26;
            }
            return sb.reverse().toString();
        }
}