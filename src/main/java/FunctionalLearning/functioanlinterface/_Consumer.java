package FunctionalLearning.functioanlinterface;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

public class _Consumer {
    public static void main(String[] args) {
        Customer nino = new Customer ("nino", "062531");
        greetCustomer (nino);
        greetCustomer.accept (nino);
        greetCustomerv2.accept (nino, true);
    }

    static BiConsumer<Customer, Boolean> greetCustomerv2 = (customer, showPhoneNumber) ->
            System.out.println ("Hello "
                    + customer.customerName + " Your Number "+ (showPhoneNumber ? customer.customerPhoneNumber:"******"));

    static Consumer<Customer> greetCustomer = customer ->
            System.out.println ("Hello "
                    + customer.customerName + " Your Number "+ customer.customerPhoneNumber);

    static void greetCustomer(Customer customer){
        System.out.println ("Hello "
        + customer.customerName + " Your Number "+ customer.customerPhoneNumber);
    }

    static class Customer {
        private final String customerName;
        private final String customerPhoneNumber;

        Customer(String customerName, String customerPhoneNumber) {
            this.customerName = customerName;
            this.customerPhoneNumber = customerPhoneNumber;
        }
    }
}
