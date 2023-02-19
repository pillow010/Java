package z_FunctionalLearning.functioanlinterface;

import java.util.function.Predicate;

public class _predicate {
    public static void main(String[] args) {
        System.out.println (isPhoneNumberValid ("070000000000"));
        System.out.println (isPhoneNumberValid ("070000000001"));
        System.out.println (isPhoneNumberValid ("170000000002"));
        System.out.println (isPhoneNumberValidPredicate.test ("070000000000"));
        System.out.println (isPhoneNumberValidPredicate.test ("070000000001"));
        System.out.println (isPhoneNumberValidPredicate.test ("070000000002"));
        System.out.println (isPhoneNumberValidPredicate.test ("170000000000"));
        System.out.println ();
        System.out.println (isPhoneNumberValidPredicate.and (containsNumber2)
                .test ("070000000002"));
        

    }

    static Boolean isPhoneNumberValid(String phoneNumber){
        return phoneNumber.startsWith ("07") && phoneNumber.length ()==12;
    }

    static Predicate<String> isPhoneNumberValidPredicate  = phoneNumber -> phoneNumber
            .startsWith ("07") && phoneNumber.length ()==12;

    static Predicate<String> containsNumber2 = phoneNumber->
            phoneNumber.contains ("2");
}
